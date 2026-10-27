// =============================================================================
//  ImageZoomAddin - 让 Word / Excel 拥有 WPS 那样的「图片放大预览」
//
//  交互：单击文档里的图片 -> 图片右下角浮出圆形按钮（放大镜；Word 还有裁剪）
//        点放大镜 -> 打开原图尺寸预览窗（滚轮缩放、拖动平移、双击 1:1/自适应、Esc 关闭）
//        点裁剪   -> 唤起 Word 原生裁剪状态（仅 Word）
//
//  技术：单个 .NET Framework 4.x 程序集，同时注册给 Word 与 Excel
//        - 不使用 VSTO（不依赖 VSTO 运行时）
//        - 不引用 Office PIA（全部后期绑定 dynamic），因此不需要任何 Office SDK
//        - 免管理员：COM 注册写在 HKCU\Software\Classes
//
//  编译：build.bat        安装：install.bat        卸载：uninstall.bat
//  日志：%LOCALAPPDATA%\ImageZoomAddin\addin.log
// =============================================================================

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using Extensibility;      // 官方 PIA：提供 IDTExtensibility2 / ext_ConnectMode / ext_DisconnectMode
using System.Windows.Forms;
using System.Xml.Linq;

// 重要：必须显式声明程序集版本。
// csc 在没有 AssemblyVersion 特性时默认编译出 0.0.0.0，
// 而 install.bat 若按 1.0.0.0 注册，mscoree 会以 0x80070002(找不到指定文件) 失败，
// Office 会因此判定加载失败并把 LoadBehavior 降级为 2（静默禁用）。
[assembly: AssemblyTitle("ImageZoomAddin")]
[assembly: AssemblyDescription("WPS-like picture magnifier preview add-in for Word and Excel")]
[assembly: AssemblyProduct("ImageZoomAddin")]
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]

namespace ImageZoomAddin
{
    // 注意：这里**不再**手写 IDTExtensibility2，而是直接引用微软官方互操作程序集
    // Extensibility.dll（GAC 里就有：Extensibility, Version=7.0.3300.0,
    // PublicKeyToken=b03f5f7f11d50a3a），用 Extensibility.IDTExtensibility2。
    //
    // 为什么必须用官方 PIA（这是踩过 4 次坑后才定位到的）：
    //   IDL 里 custom 是 [in,out] SAFEARRAY(VARIANT) *custom，托管侧对应 `ref Array`。
    //   System.Array 是抽象类型、不带元素类型信息，CLR 的**默认封送**处理不了它。
    //   官方 PIA 带 [ImportedFromTypeLib]，CLR 会改用类型库里的 ITypeInfo 规则封送，
    //   所以同一句 `ref Array` 在 PIA 里没问题。手写接口没有 TLB 来源，
    //   只能走默认封送 -> IDispatch::Invoke 返回 E_INVALIDARG 且 puArgErr=0，
    //   在 Office 里就是 OnConnection 还没进方法体、宿主以 0x80131506 崩掉。
    //   （实测：官方 PIA 的 custom 参数也是纯 Array，没有任何 MarshalAs 特性。）

    #region 日志

    internal static class Log
    {
        private static readonly object Gate = new object();
        private static string _path;

        private static string LogPath
        {
            get
            {
                if (_path == null)
                {
                    try
                    {
                        string dir = Path.Combine(
                            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                            "ImageZoomAddin");
                        if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                        _path = Path.Combine(dir, "addin.log");
                    }
                    catch
                    {
                        _path = Path.Combine(Path.GetTempPath(), "ImageZoomAddin.log");
                    }
                }
                return _path;
            }
        }

        public static void Write(string message)
        {
            try
            {
                lock (Gate)
                {
                    // 超过 1MB 就轮转，避免无限增长
                    FileInfo fi = new FileInfo(LogPath);
                    if (fi.Exists && fi.Length > 1024 * 1024)
                    {
                        try { File.Delete(LogPath + ".old"); } catch { }
                        try { File.Move(LogPath, LogPath + ".old"); } catch { }
                    }
                    File.AppendAllText(LogPath,
                        DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + "  " + message + Environment.NewLine,
                        Encoding.UTF8);
                }
            }
            catch { }
        }

        public static void Write(string message, Exception ex)
        {
            Write(message + " :: " + (ex == null ? "(null)" : ex.GetType().Name + ": " + ex.Message));
        }
    }

    #endregion

    #region Win32

    internal static class Native
    {
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll")]
        private static extern IntPtr GetDC(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);

        [DllImport("gdi32.dll")]
        private static extern int GetDeviceCaps(IntPtr hDC, int index);

        private const int LOGPIXELSX = 88;
        private const int LOGPIXELSY = 90;

        private static int _dpiX;
        private static int _dpiY;

        /// <summary>屏幕 DPI（本进程的坐标系按这个 DPI 来，DPI 不感知时 Win 会虚拟化成 96）。</summary>
        public static void GetDpi(out int dpiX, out int dpiY)
        {
            if (_dpiX <= 0 || _dpiY <= 0)
            {
                try
                {
                    IntPtr dc = GetDC(IntPtr.Zero);
                    try
                    {
                        _dpiX = GetDeviceCaps(dc, LOGPIXELSX);
                        _dpiY = GetDeviceCaps(dc, LOGPIXELSY);
                    }
                    finally { ReleaseDC(IntPtr.Zero, dc); }
                }
                catch { }
                if (_dpiX <= 0) _dpiX = 96;
                if (_dpiY <= 0) _dpiY = 96;
            }
            dpiX = _dpiX;
            dpiY = _dpiY;
        }

        // ---- 分层窗口（WS_EX_LAYERED + UpdateLayeredWindow）----
        // 用途：浮动小按钮要做"圆角 + 抗锯齿"，而 WinForms 的 Region 只是 1 位掩码，
        //       圆角边缘永远是一格一格的阶梯。走分层窗口才能真正得到 alpha 平滑的边。

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X, Y;
            public POINT(int x, int y) { X = x; Y = y; }
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct SIZE
        {
            public int cx, cy;
            public SIZE(int w, int h) { cx = w; cy = h; }
        }

        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        public struct BLENDFUNCTION
        {
            public byte BlendOp, BlendFlags, SourceConstantAlpha, AlphaFormat;
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UpdateLayeredWindow(IntPtr hwnd, IntPtr hdcDst, ref POINT pptDst,
            ref SIZE psize, IntPtr hdcSrc, ref POINT pptSrc, int crKey, ref BLENDFUNCTION pblend, int dwFlags);

        [DllImport("gdi32.dll")]
        private static extern IntPtr CreateCompatibleDC(IntPtr hdc);

        [DllImport("gdi32.dll")]
        private static extern bool DeleteDC(IntPtr hdc);

        [DllImport("gdi32.dll")]
        private static extern IntPtr SelectObject(IntPtr hdc, IntPtr obj);

        [DllImport("gdi32.dll")]
        private static extern bool DeleteObject(IntPtr obj);

        private const int ULW_ALPHA = 0x00000002;
        private const byte AC_SRC_OVER = 0x00;
        private const byte AC_SRC_ALPHA = 0x01;

        /// <summary>
        /// 把一张 32bppPArgb 位图作为窗口内容（带逐像素 alpha），并把窗口挪到 (x, y)。
        /// 返回 false 表示系统拒绝（极少见，通常是桌面合成被关掉），调用方可据此退回普通绘制。
        /// </summary>
        public static bool SetLayeredContent(IntPtr hwnd, Bitmap content, int x, int y)
        {
            if (hwnd == IntPtr.Zero || content == null) return false;

            IntPtr screenDc = IntPtr.Zero, memDc = IntPtr.Zero, hBmp = IntPtr.Zero, oldBmp = IntPtr.Zero;
            try
            {
                screenDc = GetDC(IntPtr.Zero);
                memDc = CreateCompatibleDC(screenDc);
                hBmp = content.GetHbitmap(Color.FromArgb(0));  // 0 alpha 的"背景色"，即保留位图自身 alpha
                oldBmp = SelectObject(memDc, hBmp);

                POINT src = new POINT(0, 0);
                POINT dst = new POINT(x, y);
                SIZE size = new SIZE(content.Width, content.Height);
                BLENDFUNCTION blend = new BLENDFUNCTION();
                blend.BlendOp = AC_SRC_OVER;
                blend.BlendFlags = 0;
                blend.SourceConstantAlpha = 255;
                blend.AlphaFormat = AC_SRC_ALPHA;

                return UpdateLayeredWindow(hwnd, screenDc, ref dst, ref size, memDc, ref src,
                                           0, ref blend, ULW_ALPHA);
            }
            catch { return false; }
            finally
            {
                if (oldBmp != IntPtr.Zero) SelectObject(memDc, oldBmp);
                if (hBmp != IntPtr.Zero) DeleteObject(hBmp);
                if (memDc != IntPtr.Zero) DeleteDC(memDc);
                if (screenDc != IntPtr.Zero) ReleaseDC(IntPtr.Zero, screenDc);
            }
        }

        /// <summary>当前前台窗口是否属于本进程（也就是 Word/Excel 自己）。</summary>
        public static bool IsOwnProcessForeground()
        {
#if IZ_TEST_NOFG
            // 仅用于离线自动化验证：用 /define:IZ_TEST_NOFG 编译时跳过前台判断，
            // 便于在没有真实鼠标交互的情况下检查"选中图片 -> 命中检测"这条链路。
            // 正式构建绝不定义此符号。
            return true;
#endif
            try
            {
                IntPtr h = GetForegroundWindow();
                if (h == IntPtr.Zero) return false;
                uint pid;
                GetWindowThreadProcessId(h, out pid);
                return pid == (uint)System.Diagnostics.Process.GetCurrentProcess().Id;
            }
            catch { return false; }
        }
    }

    #endregion

    #region 数据载体

    internal sealed class ImagePayload
    {
        public byte[] Bytes;
        public string Source;

        public ImagePayload(byte[] bytes, string source)
        {
            Bytes = bytes;
            Source = source;
        }
    }

    /// <summary>描述"当前选中的那张图片"。</summary>
    internal sealed class PictureHit
    {
        public string Host = "";        // Word / Excel / PowerPoint
        public bool IsInline;           // 仅 Word 有意义
        public bool IsShapeImage;       // 仅 Excel：本加载项生成的「图片填充自选图形」
        public dynamic Inline;          // Word.InlineShape
        public dynamic Shape;           // Word.Shape / Excel.Shape / PowerPoint.Shape
        public string Name = "";
        public string SheetName = "";
        public int ShapeId;             // 仅 PowerPoint：对应 OOXML 里 p:cNvPr/@id
        public int SlideIndex;          // 仅 PowerPoint
        public string Key = "";         // 用于判断"还是同一张图"
        public Rectangle ScreenRect = Rectangle.Empty;
    }

    #endregion

    #region 剪贴板：优先取原始字节（PNG/JFIF 流），其次取渲染位图

    internal static class ClipboardHelper
    {
        private static readonly string[] RawFormats = new string[]
        {
            "PNG", "image/png", "JFIF", "image/jpeg", "JPEG", "GIF", "image/gif", "TIFF"
        };

        public static byte[] TryRawImageBytes()
        {
            try
            {
                IDataObject data = Clipboard.GetDataObject();
                if (data == null) return null;

                foreach (string fmt in RawFormats)
                {
                    bool present = false;
                    try { present = data.GetDataPresent(fmt, true); } catch { }
                    if (!present) continue;

                    object raw = null;
                    try { raw = data.GetData(fmt, true); } catch { continue; }
                    Stream stream = raw as Stream;
                    if (stream == null) continue;

                    try
                    {
                        if (stream.CanSeek) stream.Position = 0;
                        byte[] bytes = ReadAll(stream);
                        if (bytes != null && bytes.Length > 128 && LooksLikeImage(bytes)) return bytes;
                    }
                    catch { }
                }
            }
            catch (Exception ex) { Log.Write("读取剪贴板原始字节失败", ex); }
            return null;
        }

        public static Bitmap TryRenderedImage()
        {
            try
            {
                if (!Clipboard.ContainsImage()) return null;
                Image img = Clipboard.GetImage();
                if (img == null) return null;
                return new Bitmap(img);
            }
            catch (Exception ex) { Log.Write("读取剪贴板位图失败", ex); return null; }
        }

        public static byte[] ImageToPngBytes(Image img)
        {
            try
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    img.Save(ms, ImageFormat.Png);
                    return ms.ToArray();
                }
            }
            catch { return null; }
        }

        private static byte[] ReadAll(Stream s)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                byte[] buffer = new byte[64 * 1024];
                int read;
                while ((read = s.Read(buffer, 0, buffer.Length)) > 0) ms.Write(buffer, 0, read);
                return ms.ToArray();
            }
        }

        internal static bool LooksLikeImage(byte[] b)
        {
            if (b == null || b.Length < 8) return false;
            if (b[0] == 0x89 && b[1] == 0x50 && b[2] == 0x4E && b[3] == 0x47) return true;      // PNG
            if (b[0] == 0xFF && b[1] == 0xD8) return true;                                        // JPEG
            if (b[0] == 0x47 && b[1] == 0x49 && b[2] == 0x46) return true;                        // GIF
            if (b[0] == 0x42 && b[1] == 0x4D) return true;                                        // BMP
            if (b[0] == 0x49 && b[1] == 0x49 && b[2] == 0x2A) return true;                        // TIFF LE
            if (b[0] == 0x4D && b[1] == 0x4D && b[2] == 0x00) return true;                        // TIFF BE
            return false;
        }
    }

    #endregion

    #region OOXML 解析：从 Flat OPC / xlsx 包里挖原始图片字节

    internal static class Ooxml
    {
        private static readonly Regex EmbedRx =
            new Regex("(?:r:embed|r:link)\\s*=\\s*\"(?<id>[^\"]+)\"", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        // ---------------------------------------------------------------- Word

        /// <summary>
        /// Word 的 Range.WordOpenXML 会返回 Flat OPC 包，图片以 pkg:binaryData(base64) 形式内嵌。
        /// 这是"原图字节"，与文档中把图片缩小到什么尺寸无关。
        /// </summary>
        public static byte[] FromFlatOpc(string flatOpc)
        {
            if (string.IsNullOrEmpty(flatOpc) || flatOpc.Length < 64) return null;

            XDocument doc;
            try { doc = XDocument.Parse(flatOpc); }
            catch (Exception ex) { Log.Write("Flat OPC 解析失败", ex); return null; }

            List<XElement> parts = doc.Descendants()
                .Where(e => e.Name.LocalName == "part")
                .ToList();
            if (parts.Count == 0) return null;

            List<XElement> imageParts = parts
                .Where(p => Attr(p, "contentType").StartsWith("image/", StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (imageParts.Count == 0) return null;
            if (imageParts.Count == 1) return PartBytes(imageParts[0]);

            // 多个图片部件：用 document 部件里的 r:embed -> rels -> target 精确定位
            try
            {
                XElement main = parts.FirstOrDefault(p =>
                {
                    string ct = Attr(p, "contentType");
                    return ct.IndexOf("wordprocessingml.document.main", StringComparison.OrdinalIgnoreCase) >= 0;
                });
                if (main == null) main = parts.FirstOrDefault(p => Attr(p, "name").EndsWith("/document.xml", StringComparison.OrdinalIgnoreCase));

                if (main != null)
                {
                    Match m = EmbedRx.Match(main.ToString());
                    if (m.Success)
                    {
                        string rid = m.Groups["id"].Value;

                        XElement rels = parts.FirstOrDefault(p =>
                        {
                            string n = Attr(p, "name");
                            return n.EndsWith(".rels", StringComparison.OrdinalIgnoreCase) &&
                                   n.IndexOf("document.xml.rels", StringComparison.OrdinalIgnoreCase) >= 0;
                        });

                        if (rels != null)
                        {
                            string target = FindRelationshipTarget(rels, rid);
                            if (target != null)
                            {
                                string wanted = NormalizePartName("/word/", target);
                                XElement hit = imageParts.FirstOrDefault(p =>
                                    string.Equals(Attr(p, "name").TrimStart('/'), wanted.TrimStart('/'),
                                                  StringComparison.OrdinalIgnoreCase));
                                if (hit != null) return PartBytes(hit);
                            }
                        }
                    }
                }
            }
            catch (Exception ex) { Log.Write("Flat OPC 定位图片失败", ex); }

            return PartBytes(imageParts[0]);
        }

        // -------------------------------------------------------------- Excel

        /// <summary>
        /// Excel 的图片存在 xl/media/*，通过 xl/drawings/drawingN.xml 里的 cNvPr/@name
        /// 与 Shape.Name 对应，再经 r:embed -> rels -> target 找到真实文件。
        /// </summary>
        public static byte[] FromExcelPackage(string xlsxPath, string shapeName, string sheetName)
        {
            if (!File.Exists(xlsxPath)) return null;

            using (ZipArchive zip = ZipFile.OpenRead(xlsxPath))
            {
                List<ZipArchiveEntry> drawings = zip.Entries
                    .Where(e => e.FullName.StartsWith("xl/drawings/", StringComparison.OrdinalIgnoreCase) &&
                                e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                    .OrderBy(e => e.FullName)
                    .ToList();

                // 先只看当前工作表的 drawing，命中率更高；再退化到全表扫描
                List<ZipArchiveEntry> ordered = drawings;
                if (!string.IsNullOrEmpty(sheetName))
                {
                    List<ZipArchiveEntry> preferred = drawings
                        .Where(e => DrawingBelongsToSheet(zip, e.FullName, sheetName))
                        .ToList();
                    if (preferred.Count > 0) ordered = preferred;
                }

                foreach (ZipArchiveEntry drawing in ordered)
                {
                    string xml = ReadText(drawing);
                    if (xml == null) continue;

                    string rid = FindEmbedForShape(xml, shapeName);
                    if (rid == null) continue;

                    string relsPath = "xl/drawings/_rels/" + Path.GetFileName(drawing.FullName) + ".rels";
                    ZipArchiveEntry rels = zip.GetEntry(relsPath);
                    if (rels == null) continue;

                    string target = FindRelationshipTargetXml(ReadText(rels), rid);
                    if (target == null) continue;

                    string mediaName = NormalizePartName("xl/drawings/", target);
                    ZipArchiveEntry media = zip.GetEntry(mediaName);
                    if (media == null)
                    {
                        media = zip.Entries.FirstOrDefault(e =>
                            string.Equals(e.FullName, mediaName, StringComparison.OrdinalIgnoreCase));
                    }
                    if (media == null) continue;

                    using (Stream s = media.Open())
                    {
                        byte[] bytes = ReadAll(s);
                        if (bytes != null && bytes.Length > 128) return bytes;
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// 从 .pptx 包里取某个形状的原始图片字节。
        ///
        /// PowerPoint 这边不能用名字定位：图片名字可以重名、也可以被用户改掉。
        /// 但 <c>p:cNvPr/@id</c> 恒等于 <c>Shape.Id</c>（实测：AddPicture 后 Shape.Id=3，
        /// 解包后 slide 里 cNvPr 的 id 也是 3），拿它当键最稳。
        ///
        /// slide 部件名优先按 SlideIndex 猜（ppt/slides/slide&lt;n&gt;.xml），猜不中再全量扫 ——
        /// 因为删过幻灯片之后，部件编号和 SlideIndex 不一定对得上。
        /// </summary>
        public static byte[] FromPowerPointPackage(string pptxPath, int shapeId, int slideIndex)
        {
            if (!File.Exists(pptxPath)) return null;
            if (shapeId <= 0) return null;

            using (ZipArchive zip = ZipFile.OpenRead(pptxPath))
            {
                List<string> candidates = new List<string>();

                string guess = "ppt/slides/slide" + slideIndex + ".xml";
                if (zip.GetEntry(guess) != null) candidates.Add(guess);

                foreach (ZipArchiveEntry e in zip.Entries.OrderBy(x => x.FullName))
                {
                    string n = e.FullName;
                    if (!n.StartsWith("ppt/slides/", StringComparison.OrdinalIgnoreCase)) continue;
                    if (!n.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)) continue;
                    if (n.IndexOf("_rels", StringComparison.OrdinalIgnoreCase) >= 0) continue;
                    if (candidates.Contains(n)) continue;
                    candidates.Add(n);
                }

                foreach (string part in candidates)
                {
                    ZipArchiveEntry slide = zip.GetEntry(part);
                    if (slide == null) continue;

                    string rid = FindEmbedForShapeId(ReadText(slide), shapeId);
                    if (rid == null) continue;

                    string relsPath = "ppt/slides/_rels/" + Path.GetFileName(part) + ".rels";
                    ZipArchiveEntry rels = zip.GetEntry(relsPath);
                    if (rels == null) continue;

                    string target = FindRelationshipTargetXml(ReadText(rels), rid);
                    if (target == null) continue;

                    string mediaName = NormalizePartName("ppt/slides/", target);
                    ZipArchiveEntry media = zip.GetEntry(mediaName);
                    if (media == null)
                    {
                        media = zip.Entries.FirstOrDefault(e =>
                            string.Equals(e.FullName, mediaName, StringComparison.OrdinalIgnoreCase));
                    }
                    if (media == null) continue;

                    using (Stream s = media.Open())
                    {
                        byte[] bytes = ReadAll(s);
                        if (bytes != null && bytes.Length > 128) return bytes;
                    }
                }
            }
            return null;
        }

        /// <summary>在 slide xml 里按 cNvPr/@id 找到那个形状，返回它 a:blip 的 r:embed。</summary>
        private static string FindEmbedForShapeId(string slideXml, int shapeId)
        {
            if (string.IsNullOrEmpty(slideXml)) return null;
            try
            {
                XDocument d = XDocument.Parse(slideXml);
                List<XElement> nodes = d.Descendants().ToList();

                string id = shapeId.ToString();
                XElement cnv = nodes.FirstOrDefault(e =>
                    e.Name.LocalName == "cNvPr" &&
                    string.Equals(Attr(e, "id"), id, StringComparison.Ordinal));
                if (cnv == null) return null;

                // 普通图片在 p:pic 里；占位符里的图片是 p:sp，blip 在 p:spPr 的 a:blipFill 下。
                XElement scope = cnv.Ancestors().FirstOrDefault(a => a.Name.LocalName == "pic");
                if (scope == null) scope = cnv.Ancestors().FirstOrDefault(a => a.Name.LocalName == "sp");
                if (scope == null) return null;

                XElement blip = scope.Descendants().FirstOrDefault(e => e.Name.LocalName == "blip");
                if (blip == null) return null;

                string embed = Attr(blip, "embed");
                return string.IsNullOrEmpty(embed) ? null : embed;
            }
            catch (Exception ex) { Log.Write("解析 slide xml 失败", ex); return null; }
        }

        private static bool DrawingBelongsToSheet(ZipArchive zip, string drawingPath, string sheetName)
        {
            // 工作表 rels 里 Target 指向 drawing，这里做反向匹配：sheetName 对应的 sheetN.xml -> rels -> drawing
            try
            {
                foreach (ZipArchiveEntry sheet in zip.Entries)
                {
                    string n = sheet.FullName;
                    if (!n.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase)) continue;
                    if (!n.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)) continue;

                    string sheetXml = ReadText(sheet);
                    if (sheetXml == null) continue;
                    // <sheet name="Sheet1" sheetId="1" r:id="rId1"/> 在 workbook.xml，这里只做弱匹配
                    if (sheetXml.IndexOf("drawing", StringComparison.OrdinalIgnoreCase) < 0) continue;

                    string rels = "xl/worksheets/_rels/" + Path.GetFileName(n) + ".rels";
                    ZipArchiveEntry relsEntry = zip.GetEntry(rels);
                    if (relsEntry == null) continue;
                    string relsXml = ReadText(relsEntry);
                    if (relsXml == null) continue;

                    string target = null;
                    try
                    {
                        XDocument d = XDocument.Parse(relsXml);
                        List<XElement> relsList = d.Descendants().Where(e => e.Name.LocalName == "Relationship").ToList();
                        foreach (XElement r in relsList)
                        {
                            string t = Attr(r, "Target");
                            if (t != null && t.IndexOf("drawings/", StringComparison.OrdinalIgnoreCase) >= 0 &&
                                Attr(r, "Type").IndexOf("/drawing", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                target = NormalizePartName("xl/worksheets/", t);
                                break;
                            }
                        }
                    }
                    catch { }

                    if (target != null && string.Equals(target, drawingPath, StringComparison.OrdinalIgnoreCase))
                        return true;
                }
            }
            catch { }
            return false;
        }

        private static string FindEmbedForShape(string drawingXml, string shapeName)
        {
            try
            {
                XDocument d = XDocument.Parse(drawingXml);
                List<XElement> nodes = d.Descendants().ToList();
                XElement cnv = nodes.FirstOrDefault(e =>
                    e.Name.LocalName == "cNvPr" &&
                    string.Equals(Attr(e, "name"), shapeName, StringComparison.OrdinalIgnoreCase));
                if (cnv == null) return null;

                // 普通图片在 xdr:pic 里；「图片填充的自选图形」是 xdr:sp + a:blipFill。
                // 少了 sp 这一层，画布上有多张图时会拿到别人第一个 blip，取到错的图。
                XElement scope = cnv.Ancestors().FirstOrDefault(a => a.Name.LocalName == "pic");
                if (scope == null) scope = cnv.Ancestors().FirstOrDefault(a => a.Name.LocalName == "sp");
                if (scope == null) scope = d.Root;
                if (scope == null) return null;

                XElement blip = scope.Descendants().FirstOrDefault(e => e.Name.LocalName == "blip");
                if (blip == null) return null;
                string embed = Attr(blip, "embed");
                if (string.IsNullOrEmpty(embed)) embed = Attr(blip, "link");
                return string.IsNullOrEmpty(embed) ? null : embed;
            }
            catch (Exception ex) { Log.Write("解析 drawing xml 失败", ex); return null; }
        }

        private static string FindRelationshipTarget(XElement relsPart, string rid)
        {
            try
            {
                foreach (XElement r in relsPart.DescendantsAndSelf().Where(e => e.Name.LocalName == "Relationship"))
                {
                    if (string.Equals(Attr(r, "Id"), rid, StringComparison.OrdinalIgnoreCase))
                        return Attr(r, "Target");
                }
            }
            catch { }
            return null;
        }

        private static string FindRelationshipTargetXml(string relsXml, string rid)
        {
            if (relsXml == null) return null;
            try
            {
                XDocument d = XDocument.Parse(relsXml);
                foreach (XElement r in d.Descendants().Where(e => e.Name.LocalName == "Relationship"))
                {
                    if (string.Equals(Attr(r, "Id"), rid, StringComparison.OrdinalIgnoreCase))
                        return Attr(r, "Target");
                }
            }
            catch (Exception ex) { Log.Write("解析 rels 失败", ex); }
            return null;
        }

        // ------------------------------------------------------------ helpers

        private static string Attr(XElement e, string localName)
        {
            XAttribute a = e.Attributes().FirstOrDefault(x => x.Name.LocalName == localName);
            return a == null ? "" : a.Value;
        }

        private static byte[] PartBytes(XElement part)
        {
            XElement bin = part.Elements().FirstOrDefault(e => e.Name.LocalName == "binaryData");
            if (bin == null) return null;
            string text = bin.Value.Trim();
            if (text.Length == 0) return null;
            try { return Convert.FromBase64String(text); }
            catch (Exception ex) { Log.Write("base64 解码失败", ex); return null; }
        }

        private static string ReadText(ZipArchiveEntry entry)
        {
            try
            {
                using (Stream s = entry.Open())
                {
                    using (StreamReader reader = new StreamReader(s, Encoding.UTF8, true))
                        return reader.ReadToEnd();
                }
            }
            catch (Exception ex) { Log.Write("读取包内条目失败: " + entry.FullName, ex); return null; }
        }

        private static byte[] ReadAll(Stream s)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                byte[] buffer = new byte[64 * 1024];
                int read;
                while ((read = s.Read(buffer, 0, buffer.Length)) > 0) ms.Write(buffer, 0, read);
                return ms.ToArray();
            }
        }

        /// <summary>把 target("../media/image1.png") 拼到 base("xl/drawings/") 上并消掉 ../</summary>
        private static string NormalizePartName(string baseDir, string target)
        {
            if (string.IsNullOrEmpty(target)) return "";
            string combined = target.StartsWith("/") ? target.TrimStart('/') : baseDir + target;

            List<string> stack = new List<string>();
            foreach (string seg in combined.Replace('\\', '/').Split('/'))
            {
                if (seg.Length == 0 || seg == ".") continue;
                if (seg == "..")
                {
                    if (stack.Count > 0) stack.RemoveAt(stack.Count - 1);
                }
                else stack.Add(seg);
            }
            return string.Join("/", stack.ToArray());
        }
    }

    #endregion

    #region 浮动的放大镜按钮

    /// <summary>三种浮动按钮：放大镜 / 裁剪 / 旋转。</summary>
    internal enum ButtonKind { Zoom, Crop, Rotate }

    /// <summary>图片右下角的圆角方框小按钮（WPS 风格：白底圆角矩形 + 线条图标）：
    /// 放大镜（预览原图）、裁剪（裁剪二级菜单）、旋转（旋转/翻转菜单）。</summary>
    internal sealed class MagnifierButton : Form
    {
        public event EventHandler Clicked;

        private readonly ButtonKind _kind;
        private bool _hover;
        private ToolTip _tip;

        public MagnifierButton() : this(ButtonKind.Zoom) { }

        public MagnifierButton(bool crop) : this(crop ? ButtonKind.Crop : ButtonKind.Zoom) { }

        // ---- WPS 风格配色（白底圆角方框 + 深灰线条 + 蓝色点缀）----
        internal const int BoxSize = 30;      // 圆角方框边长
        private const int CornerR = 6;        // 圆角半径（WPS 的圆角约为边长的 20%，不是四分之一圆）
        private const float IconScale = 1.0f;  // 图标相对 16pt 设计稿的缩放（WPS 的图标约占按钮宽度一半）

        /// <summary>超采样倍数：先按 8 倍尺寸画，再缩回 30px，圆角和细线条才有"高清"的平滑边。</summary>
        internal const int SuperSample = 8;

        private static readonly Color BodyColor = Color.White;
        private static readonly Color BodyHover = Color.FromArgb(224, 235, 251);    // 悬停底色（比之前深一档）
        private static readonly Color BorderColor = Color.FromArgb(222, 226, 234);
        private static readonly Color BorderHover = Color.FromArgb(116, 152, 214);  // 悬停描边（比之前深一档）
        private static readonly Color InkDark = Color.FromArgb(58, 58, 58);
        private static readonly Color InkDarkHover = Color.FromArgb(16, 16, 16);    // 悬停时图标压深
        private static readonly Color InkBlue = Color.FromArgb(52, 110, 212);
        private static readonly Color InkBlueHover = Color.FromArgb(24, 76, 170);   // 悬停时彩色部分压深

        private bool _rendering;
        private bool _layeredOk;

        public MagnifierButton(ButtonKind kind)
        {
            _kind = kind;

            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            TopMost = true;
            StartPosition = FormStartPosition.Manual;
            Size = new Size(BoxSize, BoxSize);
            BackColor = BodyColor;
            DoubleBuffered = true;
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);

            _tip = new ToolTip();
            _tip.SetToolTip(this, kind == ButtonKind.Crop ? "裁剪 / 形状 / 比例"
                               : kind == ButtonKind.Rotate ? "旋转 / 翻转"
                               : "图片放大预览（点击查看原图）");
            _tip.InitialDelay = 300;
        }

        protected override bool ShowWithoutActivation { get { return true; } }

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= 0x08000000;  // WS_EX_NOACTIVATE：点了也不抢焦点，Word/Excel 的选中状态不丢
                cp.ExStyle |= 0x00000080;  // WS_EX_TOOLWINDOW：不显示在 Alt+Tab
                cp.ExStyle |= 0x00080000;  // WS_EX_LAYERED：内容走 UpdateLayeredWindow，圆角才有平滑 alpha
                return cp;
            }
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            Render();
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            Render();
        }

        /// <summary>内容全部由 Render() 提交，WM_PAINT 不参与，避免和分层表面打架。</summary>
        protected override void OnPaint(PaintEventArgs e) { }

        protected override void OnPaintBackground(PaintEventArgs e) { }

        /// <summary>圆角矩形路径（半径自动收缩，不会超出矩形）。</summary>
        internal static GraphicsPath RoundedPath(RectangleF r)
        {
            float rad = Math.Min(CornerR, Math.Min(r.Width, r.Height) / 2f);
            GraphicsPath path = new GraphicsPath();
            if (rad <= 0.5f) { path.AddRectangle(r); return path; }

            float d = rad * 2f;
            path.AddArc(r.X, r.Y, d, d, 180, 90);
            path.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            path.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            path.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            Render();
        }

        /// <summary>外部（定位逻辑）改了位置后调一次，保证分层表面跟上窗口。</summary>
        public void Repaint() { Render(); }

        /// <summary>
        /// 把当前状态画成一张 32bppPArgb 位图（尺寸 = 按钮尺寸）。
        /// 先在 SuperSample 倍的大图上作画，再用高质量插值缩回来 —— 圆角、1px 描边、
        /// 细线条都因此带上真正的抗锯齿，而不是 Region 那种一格一格的阶梯。
        /// </summary>
        internal Bitmap Compose()
        {
            int ss = SuperSample;
            // 大图必须是 PArgb：GDI+ 画进去时就把 alpha 预乘好，后面做盒式平均才是对的
            using (Bitmap big = new Bitmap(BoxSize * ss, BoxSize * ss, PixelFormat.Format32bppPArgb))
            {
                using (Graphics g = Graphics.FromImage(big))
                {
                    g.Clear(Color.Transparent);
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                    g.CompositingQuality = CompositingQuality.HighQuality;
                    g.ScaleTransform(ss, ss);      // 之后一律用 30x30 的逻辑坐标作图
                    PaintBody(g);
                }
                return Downsample(big, ss);
            }
        }

        /// <summary>
        /// 整数倍盒式平均（每 ss×ss 个源像素取算术平均）。
        /// 不用 GDI+ 的插值：它缩小 8 倍时会漏采样，1px 的线缩回来是半透明的一团灰；
        /// 盒式平均是 8 倍整数缩小的正解，对齐好的 1px 线可以 100% 实心。
        /// </summary>
        private static Bitmap Downsample(Bitmap src, int factor)
        {
            int w = src.Width / factor, h = src.Height / factor;
            Bitmap dst = new Bitmap(w, h, PixelFormat.Format32bppPArgb);

            BitmapData sd = src.LockBits(new Rectangle(0, 0, src.Width, src.Height),
                                         ImageLockMode.ReadOnly, PixelFormat.Format32bppPArgb);
            BitmapData dd = dst.LockBits(new Rectangle(0, 0, w, h),
                                         ImageLockMode.WriteOnly, PixelFormat.Format32bppPArgb);
            try
            {
                int sstride = sd.Stride;
                byte[] sbuf = new byte[sstride * src.Height];
                System.Runtime.InteropServices.Marshal.Copy(sd.Scan0, sbuf, 0, sbuf.Length);

                byte[] dbuf = new byte[dd.Stride * h];
                int n = factor * factor;

                for (int y = 0; y < h; y++)
                {
                    for (int x = 0; x < w; x++)
                    {
                        int b = 0, g = 0, r = 0, a = 0;
                        int x0 = x * factor * 4;
                        int y0 = y * factor * sstride;
                        for (int yy = 0; yy < factor; yy++)
                        {
                            int row = y0 + yy * sstride + x0;
                            for (int xx = 0; xx < factor; xx++)
                            {
                                int o = row + xx * 4;
                                b += sbuf[o];
                                g += sbuf[o + 1];
                                r += sbuf[o + 2];
                                a += sbuf[o + 3];
                            }
                        }
                        int p = y * dd.Stride + x * 4;
                        dbuf[p] = (byte)(b / n);
                        dbuf[p + 1] = (byte)(g / n);
                        dbuf[p + 2] = (byte)(r / n);
                        dbuf[p + 3] = (byte)(a / n);
                    }
                }
                System.Runtime.InteropServices.Marshal.Copy(dbuf, 0, dd.Scan0, dbuf.Length);
            }
            finally
            {
                src.UnlockBits(sd);
                dst.UnlockBits(dd);
            }
            return dst;
        }

        private void Render()
        {
            if (!IsHandleCreated || _rendering) return;
            _rendering = true;
            Bitmap content = null;
            try
            {
                content = Compose();
                _layeredOk = Native.SetLayeredContent(Handle, content, Left, Top);
                if (!_layeredOk)
                    Log.Write("magnifier: 分层绘制被拒绝，按钮可能不显示（按钮=" + _kind + "）");
            }
            catch (Exception ex) { Log.Write("magnifier: 渲染失败（按钮=" + _kind + "）", ex); }
            finally
            {
                if (content != null) content.Dispose();
                _rendering = false;
            }
        }

        /// <summary>画按钮本体（逻辑坐标 30x30）。</summary>
        private void PaintBody(Graphics g)
        {
            RectangleF body = new RectangleF(0.5f, 0.5f, BoxSize - 1f, BoxSize - 1f);

            // 白色（或悬停浅蓝）圆角底板 + 1px 描边；都是抗锯齿的，所以边缘不会有锯齿
            using (GraphicsPath path = RoundedPath(body))
            {
                using (SolidBrush b = new SolidBrush(_hover ? BodyHover : BodyColor))
                    g.FillPath(b, path);
                using (Pen p = new Pen(_hover ? BorderHover : BorderColor, 1f))
                    g.DrawPath(p, path);
            }

            DrawIcon(g, _hover ? InkDarkHover : InkDark, _hover ? InkBlueHover : InkBlue);
        }

        /// <summary>
        /// 画图标：内部按 16x16 的设计稿坐标作图，再整体缩到 1.12 倍居中
        /// （WPS 的图标约占按钮宽度一半，留白多才显得轻）。
        /// dark / blue 由调用方传，悬停时压深一档。
        /// </summary>
        private void DrawIcon(Graphics g, Color dark, Color blue)
        {
            GraphicsState st = g.Save();
            g.TranslateTransform(BoxSize / 2f, BoxSize / 2f);
            g.ScaleTransform(IconScale, IconScale);
            g.TranslateTransform(-8f, -8f);

            if (_kind == ButtonKind.Crop)
            {
                // 裁剪：⌗ 两条 L 形角标 + 一道蓝色斜线。
                // 坐标全部落在 x.5 上：这样 1px 的线正好压满一个像素，
                // 经过 8 倍超采样缩回来后还是实心深灰，而不是被稀释成浅灰。
                using (Pen p = new Pen(dark, 1f))
                {
                    p.LineJoin = LineJoin.Miter;
                    g.DrawLines(p, new PointF[] {
                        new PointF(0.5f, 2.5f), new PointF(13.5f, 2.5f), new PointF(13.5f, 15.5f)
                    });
                    g.DrawLines(p, new PointF[] {
                        new PointF(2.5f, 0.5f), new PointF(2.5f, 13.5f), new PointF(15.5f, 13.5f)
                    });
                }
                using (Pen p = new Pen(blue, 1.2f))
                {
                    p.StartCap = LineCap.Round;
                    p.EndCap = LineCap.Round;
                    g.DrawLine(p, 4.0f, 12.0f, 12.0f, 4.0f);
                }
            }
            else if (_kind == ButtonKind.Rotate)
            {
                // 旋转：小方框 + 蓝色"从右边上行 -> 绕过去 -> 向左 -> 箭头"
                using (Pen p = new Pen(dark, 1f))
                    g.DrawRectangle(p, 2.5f, 6.5f, 8f, 8f);

                using (Pen p = new Pen(blue, 1.05f))
                {
                    p.StartCap = LineCap.Round;
                    p.EndCap = LineCap.Round;
                    p.LineJoin = LineJoin.Round;
                    using (GraphicsPath path = new GraphicsPath())
                    {
                        path.AddLine(13.5f, 12.5f, 13.5f, 7.5f);
                        path.AddArc(7.5f, 4.5f, 6f, 6f, 0f, -90f);
                        path.AddLine(10.5f, 4.5f, 7.5f, 4.5f);
                        g.DrawPath(p, path);
                    }
                    // 箭头（指向左）
                    g.DrawLines(p, new PointF[] {
                        new PointF(9.5f, 2.5f), new PointF(7.5f, 4.5f), new PointF(9.5f, 6.5f)
                    });
                }
            }
            else
            {
                // 放大镜：圆镜片 + 内部加号 + 右下短手柄
                using (Pen p = new Pen(dark, 1f))
                    g.DrawEllipse(p, 2.5f, 2.5f, 9f, 9f);
                using (Pen p = new Pen(dark, 1.3f))
                {
                    p.StartCap = LineCap.Round;
                    p.EndCap = LineCap.Round;
                    g.DrawLine(p, 10.6f, 10.6f, 13.4f, 13.4f);
                }
                using (Pen p = new Pen(dark, 1f))
                {
                    g.DrawLine(p, 5.5f, 7.5f, 9.5f, 7.5f);
                    g.DrawLine(p, 7.5f, 5.5f, 7.5f, 9.5f);
                }
            }

            g.Restore(st);
        }

        protected override void OnMouseEnter(EventArgs e)
        {
            _hover = true;
            Render();
            base.OnMouseEnter(e);
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            _hover = false;
            Render();
            base.OnMouseLeave(e);
        }

        protected override void OnMouseClick(MouseEventArgs e)
        {
            base.OnMouseClick(e);
            if (e.Button != MouseButtons.Left) return;
            EventHandler handler = Clicked;
            if (handler != null) handler(this, EventArgs.Empty);
        }
    }

    #endregion

    #region 裁剪面板（WPS 风格二级菜单）

    /// <summary>面板上被点中的是哪一类操作。</summary>
    internal enum CropKind { Shape, Ratio, FreeCrop, Reset }

    internal sealed class CropChoice
    {
        public CropKind Kind;
        public string Prst = "";   // 形状：OOXML 预设几何名（rect / ellipse / star5 ...）
        public double Rw, Rh;      // 比例：<= 0 表示"原图比例"
        public string Label = "";

        public static CropChoice Shape(string prst, string label)
        {
            CropChoice c = new CropChoice();
            c.Kind = CropKind.Shape;
            c.Prst = prst;
            c.Label = label;
            return c;
        }

        public static CropChoice Ratio(double rw, double rh, string label)
        {
            CropChoice c = new CropChoice();
            c.Kind = CropKind.Ratio;
            c.Rw = rw;
            c.Rh = rh;
            c.Label = label;
            return c;
        }

        public static CropChoice Other(CropKind kind, string label)
        {
            CropChoice c = new CropChoice();
            c.Kind = kind;
            c.Label = label;
            return c;
        }
    }

    /// <summary>
    /// 形状图标：把 OOXML 的预设几何（prstGeom）用 GDI+ 近似画成小线框图标。
    /// 图标只求"一眼能认出是哪个形状"，不追求与 Word 渲染像素级一致。
    /// </summary>
    internal static class ShapeIcons
    {
        public static readonly string[,] RectGroup = {
            {"rect", "矩形"}, {"roundRect", "圆角矩形"}, {"snip1Rect", "单角剪裁"},
            {"snip2SameRect", "同侧剪裁"}, {"snip2DiagRect", "对角剪裁"}, {"snipRoundRect", "单圆角"}
        };

        public static readonly string[,] BasicGroup = {
            {"ellipse", "椭圆"}, {"triangle", "等腰三角形"}, {"rtTriangle", "直角三角形"},
            {"diamond", "菱形"}, {"parallelogram", "平行四边形"}, {"trapezoid", "梯形"},
            {"pentagon", "五边形"}, {"hexagon", "六边形"}, {"octagon", "八边形"},
            {"cross", "十字形"}, {"plaque", "铭牌"}, {"chevron", "人字形"},
            {"homePlate", "箭头形"}, {"star4", "四角星"}, {"star5", "五角星"},
            {"star6", "六角星"}, {"star8", "八角星"}, {"rightArrow", "右箭头"},
            {"leftArrow", "左箭头"}, {"upArrow", "上箭头"}, {"downArrow", "下箭头"},
            {"leftRightArrow", "左右箭头"}, {"heart", "心形"}, {"cloud", "云形"},
            {"moon", "月牙"}, {"sun", "太阳"}, {"teardrop", "水滴"}, {"donut", "圆环"}
        };

        private static PointF[] Poly(params float[] xy)
        {
            PointF[] pts = new PointF[xy.Length / 2];
            for (int i = 0; i < pts.Length; i++) pts[i] = new PointF(xy[i * 2], xy[i * 2 + 1]);
            return pts;
        }

        private static PointF[] Reg(RectangleF r, int n, double startDeg)
        {
            PointF[] pts = new PointF[n];
            double cx = r.X + r.Width / 2.0, cy = r.Y + r.Height / 2.0;
            double rx = r.Width / 2.0, ry = r.Height / 2.0;
            for (int i = 0; i < n; i++)
            {
                double a = (startDeg + i * 360.0 / n) * Math.PI / 180.0;
                pts[i] = new PointF((float)(cx + rx * Math.Cos(a)), (float)(cy + ry * Math.Sin(a)));
            }
            return pts;
        }

        private static PointF[] StarPts(RectangleF r, int n, double inner)
        {
            PointF[] pts = new PointF[n * 2];
            double cx = r.X + r.Width / 2.0, cy = r.Y + r.Height / 2.0;
            double rx = r.Width / 2.0, ry = r.Height / 2.0;
            for (int i = 0; i < n * 2; i++)
            {
                double k = (i % 2 == 0) ? 1.0 : inner;
                double a = (-90.0 + i * 180.0 / n) * Math.PI / 180.0;
                pts[i] = new PointF((float)(cx + rx * k * Math.Cos(a)), (float)(cy + ry * k * Math.Sin(a)));
            }
            return pts;
        }

        /// <summary>按预设几何名生成图标路径（坐标系由 r 给出）。</summary>
        public static GraphicsPath Build(string prst, RectangleF r)
        {
            GraphicsPath gp = new GraphicsPath();
            float x = r.X, y = r.Y, w = r.Width, h = r.Height;
            float cx = x + w / 2f, cy = y + h / 2f;
            float cut = Math.Min(w, h) * 0.26f;

            switch (prst)
            {
                case "roundRect":
                    {
                        float d = Math.Min(w, h) * 0.30f;
                        gp.AddArc(x, y, d, d, 180, 90);
                        gp.AddArc(x + w - d, y, d, d, 270, 90);
                        gp.AddArc(x + w - d, y + h - d, d, d, 0, 90);
                        gp.AddArc(x, y + h - d, d, d, 90, 90);
                        gp.CloseFigure();
                        break;
                    }
                case "snip1Rect":
                    gp.AddPolygon(Poly(x, y, x + w - cut, y, x + w, y + cut, x + w, y + h, x, y + h));
                    break;
                case "snip2SameRect":
                    gp.AddPolygon(Poly(x, y, x + w - cut, y, x + w, y + cut,
                                       x + w, y + h - cut, x + w - cut, y + h, x, y + h));
                    break;
                case "snip2DiagRect":
                    gp.AddPolygon(Poly(x, y, x + w - cut, y, x + w, y + cut,
                                       x + w, y + h, x + cut, y + h, x, y + h - cut));
                    break;
                case "snipRoundRect":
                    {
                        float d = Math.Min(w, h) * 0.30f;
                        gp.AddArc(x, y, d, d, 180, 90);
                        gp.AddLine(x + w - cut, y, x + w, y + cut);
                        gp.AddLine(x + w, y + h, x, y + h);
                        gp.CloseFigure();
                        break;
                    }
                case "ellipse":
                    gp.AddEllipse(r);
                    break;
                case "triangle":
                    gp.AddPolygon(Poly(cx, y, x + w, y + h, x, y + h));
                    break;
                case "rtTriangle":
                    gp.AddPolygon(Poly(x, y, x, y + h, x + w, y + h));
                    break;
                case "diamond":
                    gp.AddPolygon(Poly(cx, y, x + w, cy, cx, y + h, x, cy));
                    break;
                case "parallelogram":
                    gp.AddPolygon(Poly(x + w * 0.24f, y, x + w, y, x + w * 0.76f, y + h, x, y + h));
                    break;
                case "trapezoid":
                    gp.AddPolygon(Poly(x + w * 0.22f, y, x + w * 0.78f, y, x + w, y + h, x, y + h));
                    break;
                case "pentagon":
                    gp.AddPolygon(Reg(r, 5, -90));
                    break;
                case "hexagon":
                    gp.AddPolygon(Reg(r, 6, 0));
                    break;
                case "octagon":
                    gp.AddPolygon(Reg(r, 8, 22.5));
                    break;
                case "cross":
                    gp.AddPolygon(Poly(
                        x + w * 0.32f, y, x + w * 0.68f, y, x + w * 0.68f, y + h * 0.32f,
                        x + w, y + h * 0.32f, x + w, y + h * 0.68f, x + w * 0.68f, y + h * 0.68f,
                        x + w * 0.68f, y + h, x + w * 0.32f, y + h, x + w * 0.32f, y + h * 0.68f,
                        x, y + h * 0.68f, x, y + h * 0.32f, x + w * 0.32f, y + h * 0.32f));
                    break;
                case "plaque":
                    {
                        float c2 = Math.Min(w, h) * 0.22f;
                        gp.AddPolygon(Poly(
                            x + c2, y, x + w - c2, y, x + w, y + c2, x + w, y + h - c2,
                            x + w - c2, y + h, x + c2, y + h, x, y + h - c2, x, y + c2));
                        break;
                    }
                case "chevron":
                    gp.AddPolygon(Poly(
                        x, y, x + w * 0.70f, y, x + w, cy, x + w * 0.70f, y + h,
                        x, y + h, x + w * 0.30f, cy));
                    break;
                case "homePlate":
                    gp.AddPolygon(Poly(x, y, x + w * 0.72f, y, x + w, cy, x + w * 0.72f, y + h, x, y + h));
                    break;
                case "star4": gp.AddPolygon(StarPts(r, 4, 0.38)); break;
                case "star5": gp.AddPolygon(StarPts(r, 5, 0.42)); break;
                case "star6": gp.AddPolygon(StarPts(r, 6, 0.55)); break;
                case "star8": gp.AddPolygon(StarPts(r, 8, 0.58)); break;
                case "rightArrow":
                    gp.AddPolygon(Poly(
                        x, y + h * 0.30f, x + w * 0.60f, y + h * 0.30f, x + w * 0.60f, y,
                        x + w, cy, x + w * 0.60f, y + h, x + w * 0.60f, y + h * 0.70f, x, y + h * 0.70f));
                    break;
                case "leftArrow":
                    gp.AddPolygon(Poly(
                        x + w, y + h * 0.30f, x + w * 0.40f, y + h * 0.30f, x + w * 0.40f, y,
                        x, cy, x + w * 0.40f, y + h, x + w * 0.40f, y + h * 0.70f, x + w, y + h * 0.70f));
                    break;
                case "upArrow":
                    gp.AddPolygon(Poly(
                        x + w * 0.30f, y + h, x + w * 0.30f, y + h * 0.40f, x, y + h * 0.40f,
                        cx, y, x + w, y + h * 0.40f, x + w * 0.70f, y + h * 0.40f, x + w * 0.70f, y + h));
                    break;
                case "downArrow":
                    gp.AddPolygon(Poly(
                        x + w * 0.30f, y, x + w * 0.30f, y + h * 0.60f, x, y + h * 0.60f,
                        cx, y + h, x + w, y + h * 0.60f, x + w * 0.70f, y + h * 0.60f, x + w * 0.70f, y));
                    break;
                case "leftRightArrow":
                    gp.AddPolygon(Poly(
                        x, cy, x + w * 0.22f, y + h * 0.20f, x + w * 0.22f, y + h * 0.34f,
                        x + w * 0.78f, y + h * 0.34f, x + w * 0.78f, y + h * 0.20f,
                        x + w, cy, x + w * 0.78f, y + h * 0.80f, x + w * 0.78f, y + h * 0.66f,
                        x + w * 0.22f, y + h * 0.66f, x + w * 0.22f, y + h * 0.80f));
                    break;
                case "heart":
                    gp.StartFigure();
                    gp.AddBezier(cx, y + h * 0.94f, x - w * 0.06f, y + h * 0.44f, x + w * 0.16f, y - h * 0.02f, cx, y + h * 0.30f);
                    gp.AddBezier(cx, y + h * 0.30f, x + w * 0.84f, y - h * 0.02f, x + w * 1.06f, y + h * 0.44f, cx, y + h * 0.94f);
                    gp.CloseFigure();
                    break;
                case "cloud":
                    gp.StartFigure();
                    gp.AddArc(x, y + h * 0.42f, w * 0.34f, h * 0.42f, 140, 190);
                    gp.AddArc(x + w * 0.14f, y + h * 0.14f, w * 0.40f, h * 0.46f, 185, 165);
                    gp.AddArc(x + w * 0.44f, y + h * 0.10f, w * 0.44f, h * 0.50f, 225, 150);
                    gp.AddArc(x + w * 0.62f, y + h * 0.42f, w * 0.34f, h * 0.42f, 290, 160);
                    gp.CloseFigure();
                    break;
                case "moon":
                    gp.StartFigure();
                    gp.AddArc(x, y, w, h, 60, 240);
                    gp.AddArc(x + w * 0.30f, y + h * 0.06f, w * 0.72f, h * 0.88f, 250, -200);
                    gp.CloseFigure();
                    break;
                case "sun":
                    gp.AddEllipse(x + w * 0.30f, y + h * 0.30f, w * 0.40f, h * 0.40f);
                    for (int i = 0; i < 8; i++)
                    {
                        double a = i * Math.PI / 4.0;
                        gp.StartFigure();
                        gp.AddLine(
                            (float)(cx + Math.Cos(a) * w * 0.40f), (float)(cy + Math.Sin(a) * h * 0.40f),
                            (float)(cx + Math.Cos(a) * w * 0.50f), (float)(cy + Math.Sin(a) * h * 0.50f));
                    }
                    break;
                case "teardrop":
                    gp.StartFigure();
                    gp.AddBezier(cx, y, x + w * 0.92f, y + h * 0.34f, x + w, y + h, cx, y + h);
                    gp.AddBezier(cx, y + h, x, y + h, x + w * 0.08f, y + h * 0.34f, cx, y);
                    gp.CloseFigure();
                    break;
                case "donut":
                    gp.AddEllipse(r);
                    gp.StartFigure();
                    gp.AddEllipse(x + w * 0.28f, y + h * 0.28f, w * 0.44f, h * 0.44f);
                    break;
                default:
                    gp.AddRectangle(r);
                    break;
            }
            return gp;
        }

        /// <summary>比例图标：在方块里按比例画一个小矩形。</summary>
        public static GraphicsPath Ratio(double rw, double rh, RectangleF box)
        {
            GraphicsPath gp = new GraphicsPath();
            if (rw <= 0 || rh <= 0)
            {
                // 原图比例：一个带箭头的方框
                gp.AddRectangle(new RectangleF(box.X, box.Y, box.Width, box.Height));
                return gp;
            }
            double k = rw / rh;
            float bw = box.Width, bh = box.Height;
            float w, h;
            if (k >= 1.0) { w = bw; h = (float)(bw / k); }
            else { h = bh; w = (float)(bh * k); }
            float x = box.X + (bw - w) / 2f, y = box.Y + (bh - h) / 2f;
            gp.AddRectangle(new RectangleF(x, y, w, h));
            return gp;
        }
    }

    /// <summary>
    /// 裁剪按钮右侧弹出的二级菜单（WPS 的样子）：
    ///   两个标签页「按形状裁剪 / 按比例裁剪」，底部「自由裁剪」「重设形状和大小」。
    /// 整块面板自绘，没有子控件 —— 子控件会各自处理鼠标、容易把焦点抢走。
    /// 窗口同样带 WS_EX_NOACTIVATE：点面板不会让 Word 失焦，图片的选中状态一直在。
    /// </summary>
    internal sealed class CropMenuForm : Form
    {
        public event Action<CropChoice> Chosen;

        /// <summary>
        /// 这块区域里的点击不算"点外面"（就是裁剪按钮本身）。
        /// 否则会出现这个竞态：点按钮 -> 面板先被定时器关掉 -> 按钮的 Click 再把它打开，
        /// 用户点一下等于没反应。
        /// </summary>
        public Rectangle KeepOpenRect = Rectangle.Empty;

        private const int PanelW = 236;
        private const int TabH = 34;
        private const int Pad = 11;
        private const int CellW = 34;
        private const int CellH = 34;
        private const int CellGap = 2;
        private const int GroupH = 22;
        private const int FooterH = 62;
        private const float IconSize = 18f;
        private const int Cols = 6;

        private static readonly Color IconColor = Color.FromArgb(52, 110, 212);
        private static readonly Color TextColor = Color.FromArgb(48, 48, 48);
        private static readonly Color MutedColor = Color.FromArgb(110, 110, 110);
        private static readonly Color AccentColor = Color.FromArgb(0, 103, 192);
        private static readonly Color HoverColor = Color.FromArgb(233, 241, 251);

        private sealed class Item
        {
            public Rectangle Rect;
            public CropChoice Choice;
        }

        private sealed class Header
        {
            public string Title;
            public Rectangle Rect;
        }

        private static readonly string[,] RatioTable = {
            {"1:1", "1", "1"}, {"4:3", "4", "3"}, {"3:4", "3", "4"},
            {"16:9", "16", "9"}, {"9:16", "9", "16"}, {"3:2", "3", "2"},
            {"2:3", "2", "3"}, {"5:4", "5", "4"}, {"4:5", "4", "5"},
            {"21:9", "21", "9"}, {"原图比例", "0", "0"}
        };

        private int _tab;                       // 0 = 按形状裁剪，1 = 按比例裁剪
        private int _scroll;
        private int _contentH;
        private int _hover = -1;                // 命中索引（含底部两行）
        private List<Item> _items = new List<Item>();
        private List<Header> _headers = new List<Header>();
        private Rectangle _tabRect0, _tabRect1;
        private Rectangle _freeRect, _resetRect;
        private Timer _closeTimer;

        public CropMenuForm()
        {
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            TopMost = true;
            StartPosition = FormStartPosition.Manual;
            BackColor = Color.White;
            DoubleBuffered = true;
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);

            Rebuild();
            UpdateRegion();
        }

        protected override bool ShowWithoutActivation { get { return true; } }

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= 0x08000000;   // WS_EX_NOACTIVATE
                cp.ExStyle |= 0x00000080;   // WS_EX_TOOLWINDOW
                return cp;
            }
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            _closeTimer = new Timer();
            _closeTimer.Interval = 120;
            _closeTimer.Tick += OnCloseTick;
            _closeTimer.Start();
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            if (_closeTimer != null)
            {
                _closeTimer.Stop();
                _closeTimer.Dispose();
                _closeTimer = null;
            }
            base.OnFormClosed(e);
        }

        /// <summary>点面板外面就自动收起（窗口是 NOACTIVATE 的，收不到 Deactivate，只能自己盯鼠标）。</summary>
        private void OnCloseTick(object sender, EventArgs e)
        {
            try
            {
                if (!Visible) return;
                if ((Control.MouseButtons & MouseButtons.Left) == 0) return;

                Point cur = Cursor.Position;
                if (Bounds.Contains(cur)) return;
                if (!KeepOpenRect.IsEmpty && KeepOpenRect.Contains(cur)) return;
                Close();
            }
            catch { }
        }

        // ---------------------------------------------------------- 布局

        private void Rebuild()
        {
            _items.Clear();
            _headers.Clear();

            int viewport = ContentViewportHeight();
            int y = 0;
            int x = Pad;

            if (_tab == 0)
            {
                y = AddGroup("矩形", y, viewport, ref x, ShapeIcons.RectGroup);
                y += 10;
                x = Pad;
                y = AddGroup("基本形状", y, viewport, ref x, ShapeIcons.BasicGroup);
                y += 6;
            }
            else
            {
                y = AddGroup("常用比例", y, viewport, ref x, null);
                int cellW = (PanelW - Pad * 2 - CellGap) / 2;
                for (int i = 0; i < RatioTable.GetLength(0); i++)
                {
                    int col = i % 2, row = i / 2;
                    double rw = Convert.ToDouble(RatioTable[i, 1], System.Globalization.CultureInfo.InvariantCulture);
                    double rh = Convert.ToDouble(RatioTable[i, 2], System.Globalization.CultureInfo.InvariantCulture);
                    Rectangle r = new Rectangle(Pad + col * (cellW + CellGap),
                                                y + row * (CellH + 4), cellW, CellH);
                    Item it = new Item();
                    it.Rect = r;
                    it.Choice = CropChoice.Ratio(rw, rh, RatioTable[i, 0]);
                    _items.Add(it);
                }
                y += ((RatioTable.GetLength(0) + 1) / 2) * (CellH + 4) + 6;
            }

            _contentH = y;
            int maxScroll = Math.Max(0, _contentH - viewport);
            if (_scroll > maxScroll) _scroll = maxScroll;
            if (_scroll < 0) _scroll = 0;

            _tabRect0 = new Rectangle(0, 0, PanelW / 2, TabH);
            _tabRect1 = new Rectangle(PanelW / 2, 0, PanelW - PanelW / 2, TabH);
            _freeRect = new Rectangle(0, Height - FooterH + 1, PanelW, 30);
            _resetRect = new Rectangle(0, Height - 31, PanelW, 30);
        }

        private int AddGroup(string title, int y, int viewport, ref int x, string[,] table)
        {
            Header hd = new Header();
            hd.Title = title;
            hd.Rect = new Rectangle(0, y, PanelW, GroupH);
            _headers.Add(hd);
            y += GroupH;

            if (table != null)
            {
                for (int i = 0; i < table.GetLength(0); i++)
                {
                    if (x + CellW > PanelW - Pad)
                    {
                        x = Pad;
                        y += CellH + CellGap;
                    }
                    Item it = new Item();
                    it.Rect = new Rectangle(x, y, CellW, CellH);
                    it.Choice = CropChoice.Shape(table[i, 0], table[i, 1]);
                    _items.Add(it);
                    x += CellW + CellGap;
                }
                y += CellH + 2;
            }
            else
            {
                y += 4;
            }
            return y;
        }

        private int ContentViewportHeight()
        {
            return Height - TabH - FooterH;
        }

        /// <summary>按内容算出面板高度，并摆到屏幕内。</summary>
        public void Relayout()
        {
            int viewport;
            if (_tab == 0)
            {
                // 矩形（22 + 一行 36 + 间距 10）+ 基本形状（22 + 5 行 36）
                int rows = (ShapeIcons.BasicGroup.GetLength(0) + Cols - 1) / Cols;
                viewport = GroupH + CellH + CellGap + 10 + GroupH + rows * (CellH + CellGap) + 6;
            }
            else
            {
                viewport = GroupH + 4 + ((RatioTable.GetLength(0) + 1) / 2) * (CellH + 4) + 6;
            }
            if (viewport > 264) viewport = 264;
            ClientSize = new Size(PanelW, TabH + viewport + FooterH);
            Rebuild();
            UpdateRegion();
        }

        private void UpdateRegion()
        {
            try
            {
                using (GraphicsPath path = Round(new RectangleF(0, 0, Width, Height), 8f))
                    Region = new Region(path);
            }
            catch { }
        }

        private static GraphicsPath Round(RectangleF r, float radius)
        {
            GraphicsPath p = new GraphicsPath();
            float d = radius * 2f;
            p.AddArc(r.X, r.Y, d, d, 180, 90);
            p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            p.CloseFigure();
            return p;
        }

        // ---------------------------------------------------------- 绘制

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
            g.Clear(Color.White);

            using (Font tabFont = new Font("Microsoft YaHei UI", 9.5f))
            using (Font groupFont = new Font("Microsoft YaHei UI", 8.5f))
            using (Font itemFont = new Font("Microsoft YaHei UI", 9f))
            using (StringFormat sf = new StringFormat())
            {
                sf.Alignment = StringAlignment.Center;
                sf.LineAlignment = StringAlignment.Center;

                // ---- 标签页
                DrawTab(g, tabFont, _tabRect0, "按形状裁剪", _tab == 0);
                DrawTab(g, tabFont, _tabRect1, "按比例裁剪", _tab == 1);
                using (Pen p = new Pen(Color.FromArgb(226, 226, 226)))
                    g.DrawLine(p, 0, TabH - 0.5f, PanelW, TabH - 0.5f);

                // ---- 内容区（可滚动）
                Rectangle content = new Rectangle(0, TabH, PanelW, ContentViewportHeight());
                GraphicsState st = g.Save();
                g.SetClip(content);
                g.TranslateTransform(0, TabH - _scroll);

                foreach (Header hd in _headers)
                {
                    using (SolidBrush b = new SolidBrush(Color.FromArgb(245, 245, 245)))
                        g.FillRectangle(b, hd.Rect);
                    using (SolidBrush b = new SolidBrush(MutedColor))
                        g.DrawString(hd.Title, groupFont, b, new PointF(12, hd.Rect.Y + hd.Rect.Height / 2f),
                                     new StringFormat { LineAlignment = StringAlignment.Center });
                }

                for (int i = 0; i < _items.Count; i++)
                {
                    if (_items[i].Rect.Bottom + TabH - _scroll < TabH) continue;
                    if (_items[i].Rect.Top + TabH - _scroll > content.Bottom) continue;
                    DrawItem(g, itemFont, _items[i], i == _hover);
                }

                g.Restore(st);

                // ---- 底部两行
                using (Pen p = new Pen(Color.FromArgb(230, 230, 230)))
                    g.DrawLine(p, 0, Height - FooterH + 0.5f, PanelW, Height - FooterH + 0.5f);

                DrawFooterRow(g, itemFont, _freeRect, "自由裁剪", false, _hover == 1000);
                DrawFooterRow(g, itemFont, _resetRect, "重设形状和大小", true, _hover == 1001);

                // ---- 边框
                using (Pen p = new Pen(Color.FromArgb(205, 205, 205)))
                using (GraphicsPath bp = Round(new RectangleF(0.5f, 0.5f, Width - 1, Height - 1), 8f))
                    g.DrawPath(p, bp);
            }
        }

        private void DrawTab(Graphics g, Font font, Rectangle r, string text, bool active)
        {
            using (StringFormat f = new StringFormat())
            {
                f.Alignment = StringAlignment.Center;
                f.LineAlignment = StringAlignment.Center;
                using (SolidBrush b = new SolidBrush(active ? AccentColor : TextColor))
                    g.DrawString(text, font, b, r, f);

                if (active)
                {
                    SizeF sz = g.MeasureString(text, font);
                    float x = r.X + (r.Width - sz.Width) / 2f;
                    using (Pen p = new Pen(AccentColor, 2f))
                        g.DrawLine(p, x, r.Bottom - 3, x + sz.Width, r.Bottom - 3);
                }
            }
        }

        private void DrawItem(Graphics g, Font font, Item it, bool hover)
        {
            if (hover)
            {
                using (GraphicsPath hp = Round(new RectangleF(it.Rect.X + 1, it.Rect.Y + 1,
                                                              it.Rect.Width - 2, it.Rect.Height - 2), 4f))
                using (SolidBrush b = new SolidBrush(HoverColor))
                    g.FillPath(b, hp);
            }

            if (it.Choice.Kind == CropKind.Shape)
            {
                RectangleF box = new RectangleF(it.Rect.X + (it.Rect.Width - IconSize) / 2f,
                                                it.Rect.Y + (it.Rect.Height - IconSize) / 2f,
                                                IconSize, IconSize);
                using (GraphicsPath p = ShapeIcons.Build(it.Choice.Prst, box))
                using (Pen pen = new Pen(IconColor, 1.35f))
                {
                    pen.LineJoin = LineJoin.Round;
                    g.DrawPath(pen, p);
                }
            }
            else
            {
                RectangleF box = new RectangleF(it.Rect.X + 8, it.Rect.Y + (it.Rect.Height - 15) / 2f, 20, 15);
                using (GraphicsPath p = ShapeIcons.Ratio(it.Choice.Rw, it.Choice.Rh, box))
                using (Pen pen = new Pen(IconColor, 1.25f))
                    g.DrawPath(pen, p);

                using (StringFormat f = new StringFormat())
                {
                    f.LineAlignment = StringAlignment.Center;
                    using (SolidBrush b = new SolidBrush(TextColor))
                        g.DrawString(it.Choice.Label, font, b, new PointF(it.Rect.X + 34, it.Rect.Y + it.Rect.Height / 2f), f);
                }
            }
        }

        private void DrawFooterRow(Graphics g, Font font, Rectangle r, string text, bool reset, bool hover)
        {
            if (hover)
            {
                using (SolidBrush b = new SolidBrush(HoverColor))
                    g.FillRectangle(b, r);
            }

            RectangleF icon = new RectangleF(r.X + 12, r.Y + (r.Height - 15) / 2f, 16, 15);
            using (Pen pen = new Pen(hover ? AccentColor : Color.FromArgb(96, 96, 96), 1.3f))
            {
                if (reset)
                {
                    // 方块 + 回转箭头
                    g.DrawRectangle(pen, icon.X, icon.Y + 3, icon.Width - 5, icon.Height - 3);
                    using (GraphicsPath p = new GraphicsPath())
                    {
                        p.AddArc(icon.X + 2, icon.Y - 1, 8, 8, 200, 250);
                        g.DrawPath(pen, p);
                    }
                    g.DrawLine(pen, icon.X + 2, icon.Y - 1.5f, icon.X + 4f, icon.Y + 2.5f);
                }
                else
                {
                    // 裁剪 ⌗
                    g.DrawLines(pen, new PointF[] {
                        new PointF(icon.X + 4, icon.Y), new PointF(icon.X + 4, icon.Y + 13),
                        new PointF(icon.X + 16, icon.Y + 13) });
                    g.DrawLines(pen, new PointF[] {
                        new PointF(icon.X, icon.Y + 4), new PointF(icon.X + 12, icon.Y + 4),
                        new PointF(icon.X + 12, icon.Y + 16) });
                }
            }

            using (StringFormat f = new StringFormat())
            {
                f.LineAlignment = StringAlignment.Center;
                using (SolidBrush b = new SolidBrush(hover ? AccentColor : TextColor))
                    g.DrawString(text, font, b, new PointF(r.X + 34, r.Y + r.Height / 2f), f);
            }
        }

        // ---------------------------------------------------------- 交互

        /// <summary>把鼠标位置换算成命中项；-1 表示空白，1000/1001 表示底部两行。</summary>
        private int HitTest(Point p)
        {
            if (_freeRect.Contains(p)) return 1000;
            if (_resetRect.Contains(p)) return 1001;

            Rectangle content = new Rectangle(0, TabH, PanelW, ContentViewportHeight());
            if (!content.Contains(p)) return -1;

            int y = p.Y - TabH + _scroll;
            for (int i = 0; i < _items.Count; i++)
            {
                Rectangle r = _items[i].Rect;
                if (r.Contains(p.X, y)) return i;
            }
            return -1;
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            int h = HitTest(e.Location);
            if (h != _hover) { _hover = h; Invalidate(); }
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            if (_hover != -1) { _hover = -1; Invalidate(); }
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            if (e.Button != MouseButtons.Left) return;

            // 标签页
            if (_tabRect0.Contains(e.Location))
            {
                if (_tab != 0) { _tab = 0; _scroll = 0; Relayout(); Invalidate(); }
                return;
            }
            if (_tabRect1.Contains(e.Location))
            {
                if (_tab != 1) { _tab = 1; _scroll = 0; Relayout(); Invalidate(); }
                return;
            }

            int h = HitTest(e.Location);
            CropChoice choice = null;
            if (h == 1000) choice = CropChoice.Other(CropKind.FreeCrop, "自由裁剪");
            else if (h == 1001) choice = CropChoice.Other(CropKind.Reset, "重设形状和大小");
            else if (h >= 0) choice = _items[h].Choice;

            if (choice == null) return;

            Action<CropChoice> handler = Chosen;
            if (handler != null) handler(choice);
        }

        protected override void OnMouseWheel(MouseEventArgs e)
        {
            base.OnMouseWheel(e);
            int viewport = ContentViewportHeight();
            int max = Math.Max(0, _contentH - viewport);
            if (max <= 0) return;
            _scroll -= e.Delta > 0 ? 36 : -36;
            if (_scroll < 0) _scroll = 0;
            if (_scroll > max) _scroll = max;
            Invalidate();
        }
    }

    #endregion

    #region 旋转面板（旋转 / 翻转，对应「图片格式 → 排列 → 旋转」）

    internal enum RotateKind { Right90, Left90, FlipH, FlipV }

    /// <summary>
    /// 点了旋转按钮后弹出来的小面板。内容就是 Office「图片格式 → 排列 → 旋转」
    /// 下拉里的四项：向右旋转 90°、向左旋转 90°、水平翻转、垂直翻转。
    /// </summary>
    internal sealed class RotateMenuForm : Form
    {
        public event Action<RotateKind> Chosen;

        /// <summary>这块区域里的点击不算"点外面"（旋转按钮本身），避免"点一下等于没反应"的竞态。</summary>
        public Rectangle KeepOpenRect = Rectangle.Empty;

        private const int PanelW = 150;
        private const int RowH = 30;
        private const int PadY = 5;

        private static readonly Color TextColor = Color.FromArgb(48, 48, 48);
        private static readonly Color HoverColor = Color.FromArgb(233, 241, 251);
        private static readonly Color IconColor = Color.FromArgb(52, 110, 212);

        private static readonly string[] Labels = {
            "向右旋转 90°", "向左旋转 90°", "水平翻转", "垂直翻转"
        };

        private static readonly RotateKind[] Kinds = {
            RotateKind.Right90, RotateKind.Left90, RotateKind.FlipH, RotateKind.FlipV
        };

        private int _hover = -1;
        private Timer _closeTimer;

        public RotateMenuForm()
        {
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            TopMost = true;
            StartPosition = FormStartPosition.Manual;
            BackColor = Color.White;
            DoubleBuffered = true;
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);

            ClientSize = new Size(PanelW, PadY * 2 + RowH * Kinds.Length);
            UpdateRegion();
        }

        protected override bool ShowWithoutActivation { get { return true; } }

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= 0x08000000;   // WS_EX_NOACTIVATE
                cp.ExStyle |= 0x00000080;   // WS_EX_TOOLWINDOW
                return cp;
            }
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            _closeTimer = new Timer();
            _closeTimer.Interval = 120;
            _closeTimer.Tick += OnCloseTick;
            _closeTimer.Start();
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            if (_closeTimer != null)
            {
                _closeTimer.Stop();
                _closeTimer.Dispose();
                _closeTimer = null;
            }
            base.OnFormClosed(e);
        }

        private void OnCloseTick(object sender, EventArgs e)
        {
            try
            {
                if (!Visible) return;
                if ((Control.MouseButtons & MouseButtons.Left) == 0) return;
                Point cur = Cursor.Position;
                if (Bounds.Contains(cur)) return;
                if (!KeepOpenRect.IsEmpty && KeepOpenRect.Contains(cur)) return;
                Close();
            }
            catch { }
        }

        private void UpdateRegion()
        {
            try
            {
                using (GraphicsPath path = new GraphicsPath())
                {
                    RectangleF r = new RectangleF(0, 0, Width, Height);
                    float d = 16f;
                    path.AddArc(r.X, r.Y, d, d, 180, 90);
                    path.AddArc(r.Right - d, r.Y, d, d, 270, 90);
                    path.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
                    path.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
                    path.CloseFigure();
                    Region = new Region(path);
                }
            }
            catch { }
        }

        private Rectangle RowRect(int i)
        {
            return new Rectangle(4, PadY + i * RowH, PanelW - 8, RowH - 2);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
            g.Clear(Color.White);

            using (Font f = new Font("Microsoft YaHei UI", 9f))
            using (StringFormat sf = new StringFormat())
            {
                sf.LineAlignment = StringAlignment.Center;
                sf.Alignment = StringAlignment.Near;

                for (int i = 0; i < Kinds.Length; i++)
                {
                    Rectangle r = RowRect(i);
                    if (i == _hover)
                        using (SolidBrush b = new SolidBrush(HoverColor))
                            g.FillRectangle(b, r);

                    DrawGlyph(g, IconColor, new Rectangle(r.X + 6, r.Y + (r.Height - 16) / 2, 16, 16), Kinds[i]);

                    using (SolidBrush b = new SolidBrush(TextColor))
                        g.DrawString(Labels[i], f, b, new RectangleF(r.X + 30, r.Y, r.Width - 32, r.Height), sf);
                }
            }

            using (Pen p = new Pen(Color.FromArgb(205, 205, 205)))
            using (GraphicsPath bp = new GraphicsPath())
            {
                RectangleF r = new RectangleF(0.5f, 0.5f, Width - 1, Height - 1);
                float d = 16f;
                bp.AddArc(r.X, r.Y, d, d, 180, 90);
                bp.AddArc(r.Right - d, r.Y, d, d, 270, 90);
                bp.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
                bp.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
                bp.CloseFigure();
                g.DrawPath(p, bp);
            }
        }

        /// <summary>每行左边的小图形：旋转用弧线箭头，翻转用带三角的对称折线。</summary>
        private static void DrawGlyph(Graphics g, Color color, Rectangle box, RotateKind kind)
        {
            float cx = box.X + box.Width / 2f;
            float cy = box.Y + box.Height / 2f;

            if (kind == RotateKind.Right90 || kind == RotateKind.Left90)
            {
                bool right = kind == RotateKind.Right90;
                using (Pen p = new Pen(color, 1.5f))
                {
                    p.StartCap = LineCap.Round;
                    p.EndCap = LineCap.Round;
                    if (right) g.DrawArc(p, cx - 6, cy - 6, 12, 12, -60, 300);
                    else g.DrawArc(p, cx - 6, cy - 6, 12, 12, 120, -300);
                }
                using (SolidBrush b = new SolidBrush(color))
                {
                    PointF tip = right ? new PointF(cx + 6.2f, cy - 3.0f) : new PointF(cx - 6.2f, cy - 3.0f);
                    PointF b1 = right ? new PointF(cx + 1.2f, cy - 2.8f) : new PointF(cx - 1.2f, cy - 2.8f);
                    PointF b2 = right ? new PointF(cx + 4.3f, cy - 8.0f) : new PointF(cx - 4.3f, cy - 8.0f);
                    g.FillPolygon(b, new PointF[] { tip, b1, b2 });
                }
                return;
            }

            // 翻转：一个实心三角 + 一个空心三角，中间虚线
            bool horizontal = kind == RotateKind.FlipH;
            using (Pen p = new Pen(color, 1.3f))
            using (SolidBrush b = new SolidBrush(color))
            {
                if (horizontal)
                {
                    g.FillPolygon(b, new PointF[] {
                        new PointF(cx - 1.5f, cy - 6), new PointF(cx - 1.5f, cy + 6), new PointF(cx - 6.5f, cy)
                    });
                    g.DrawPolygon(p, new PointF[] {
                        new PointF(cx + 1.5f, cy - 6), new PointF(cx + 1.5f, cy + 6), new PointF(cx + 6.5f, cy)
                    });
                    p.DashStyle = DashStyle.Dot;
                    g.DrawLine(p, cx, box.Y + 1, cx, box.Bottom - 1);
                }
                else
                {
                    g.FillPolygon(b, new PointF[] {
                        new PointF(cx - 6, cy - 1.5f), new PointF(cx + 6, cy - 1.5f), new PointF(cx, cy - 6.5f)
                    });
                    g.DrawPolygon(p, new PointF[] {
                        new PointF(cx - 6, cy + 1.5f), new PointF(cx + 6, cy + 1.5f), new PointF(cx, cy + 6.5f)
                    });
                    p.DashStyle = DashStyle.Dot;
                    g.DrawLine(p, box.X + 1, cy, box.Right - 1, cy);
                }
            }
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            int hover = -1;
            for (int i = 0; i < Kinds.Length; i++)
                if (RowRect(i).Contains(e.Location)) { hover = i; break; }

            if (hover != _hover) { _hover = hover; Invalidate(); }
            base.OnMouseMove(e);
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            if (_hover != -1) { _hover = -1; Invalidate(); }
            base.OnMouseLeave(e);
        }

        protected override void OnMouseClick(MouseEventArgs e)
        {
            base.OnMouseClick(e);
            if (e.Button != MouseButtons.Left) return;

            for (int i = 0; i < Kinds.Length; i++)
            {
                if (!RowRect(i).Contains(e.Location)) continue;
                Action<RotateKind> h = Chosen;
                Close();
                if (h != null) h(Kinds[i]);
                return;
            }
        }
    }

    #endregion

    #region 缩放画布（滚轮缩放 / 拖动平移 / 1:1 / 自适应）

    internal sealed class ZoomCanvas : Control
    {
        private const float MinZoom = 0.02f;
        private const float MaxZoom = 40f;

        private Image _img;
        private float _zoom = 1f;
        private float _ox;
        private float _oy;
        private bool _panning;
        private Point _panStart;
        private float _panOx, _panOy;

        public event EventHandler ZoomChanged;

        public ZoomCanvas()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.UserPaint | ControlStyles.ResizeRedraw | ControlStyles.Selectable, true);
            BackColor = Color.FromArgb(42, 42, 45);
            TabStop = true;
        }

        public float Zoom { get { return _zoom; } }
        public Image SourceImage { get { return _img; } }

        public void SetImage(Image image)
        {
            _img = image;
            InitialView();
        }

        public void InitialView()
        {
            if (_img == null) return;
            float fit = FitZoom();
            // 原图尺寸优先：能一屏放下就 1:1，放不下才自适应
            _zoom = fit < 1f ? fit : 1f;
            Center();
        }

        public void Fit()
        {
            if (_img == null) return;
            _zoom = Math.Max(FitZoom(), MinZoom);
            Center();
        }

        public void ActualSize()
        {
            if (_img == null) return;
            _zoom = 1f;
            Center();
        }

        public void ZoomBy(float factor)
        {
            ZoomAt(new Point(ClientSize.Width / 2, ClientSize.Height / 2), factor);
        }

        /// <summary>以屏幕上某点为中心缩放（该点下的像素保持不动）。</summary>
        public void ZoomAt(Point screenPoint, float factor)
        {
            if (_img == null) return;
            float target = Clamp(_zoom * factor);
            if (Math.Abs(target - _zoom) < 0.0001f) return;

            float ix = (screenPoint.X - _ox) / _zoom;
            float iy = (screenPoint.Y - _oy) / _zoom;

            _zoom = target;
            _ox = screenPoint.X - ix * _zoom;
            _oy = screenPoint.Y - iy * _zoom;

            Invalidate();
            RaiseZoom();
        }

        private float FitZoom()
        {
            if (_img == null || ClientSize.Width <= 2 || ClientSize.Height <= 2) return 1f;
            float zx = (float)(ClientSize.Width - 24) / _img.Width;
            float zy = (float)(ClientSize.Height - 24) / _img.Height;
            return Math.Min(zx, zy);
        }

        private void Center()
        {
            if (_img == null) return;
            _ox = (ClientSize.Width - _img.Width * _zoom) / 2f;
            _oy = (ClientSize.Height - _img.Height * _zoom) / 2f;
            Invalidate();
            RaiseZoom();
        }

        private static float Clamp(float z)
        {
            if (z < MinZoom) return MinZoom;
            if (z > MaxZoom) return MaxZoom;
            return z;
        }

        private void RaiseZoom()
        {
            EventHandler handler = ZoomChanged;
            if (handler != null) handler(this, EventArgs.Empty);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.Clear(BackColor);
            if (_img == null) return;

            // 缩小：高质量插值；放大到 100% 以上：最近邻，看像素细节不糊
            g.InterpolationMode = _zoom >= 1f ? InterpolationMode.NearestNeighbor : InterpolationMode.HighQualityBicubic;
            g.PixelOffsetMode = PixelOffsetMode.Half;

            GraphicsState state = g.Save();
            g.TranslateTransform(_ox, _oy);
            g.ScaleTransform(_zoom, _zoom);
            g.DrawImage(_img, 0f, 0f, _img.Width, _img.Height);
            g.Restore(state);

            DrawOverlay(g);
        }

        private void DrawOverlay(Graphics g)
        {
            try
            {
                string left = string.Format("{0:0}%   {1} × {2} px", _zoom * 100f, _img.Width, _img.Height);
                string right = "滚轮缩放 · 左键拖动 · 双击 1:1/自适应 · Esc 关闭";

                g.SmoothingMode = SmoothingMode.AntiAlias;

                using (Font f = new Font("Microsoft YaHei UI", 8.5f))
                using (Brush text = new SolidBrush(Color.FromArgb(235, 245, 245, 245)))
                using (Brush back = new SolidBrush(Color.FromArgb(150, 0, 0, 0)))
                {
                    SizeF s1 = g.MeasureString(left, f);
                    SizeF s2 = g.MeasureString(right, f);
                    int barH = (int)Math.Max(s1.Height, s2.Height) + 8;
                    int y = ClientSize.Height - barH - 8;

                    RectangleF r1 = new RectangleF(8, y, s1.Width + 16, barH);
                    RectangleF r2 = new RectangleF(ClientSize.Width - s2.Width - 24, y, s2.Width + 16, barH);

                    g.FillRectangle(back, r1);
                    g.FillRectangle(back, r2);
                    g.DrawString(left, f, text, 16, y + 4);
                    g.DrawString(right, f, text, ClientSize.Width - s2.Width - 16, y + 4);
                }
            }
            catch { }
        }

        // ---------------------------------------------------------- 交互

        protected override void OnMouseWheel(MouseEventArgs e)
        {
            base.OnMouseWheel(e);
            ZoomAt(e.Location, e.Delta > 0 ? 1.25f : 0.8f);
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            Focus();
            if (e.Button == MouseButtons.Left && _img != null)
            {
                _panning = true;
                _panStart = e.Location;
                _panOx = _ox;
                _panOy = _oy;
                Cursor = Cursors.SizeAll;
            }
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            if (!_panning) return;
            _ox = _panOx + (e.X - _panStart.X);
            _oy = _panOy + (e.Y - _panStart.Y);
            Invalidate();
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            if (e.Button == MouseButtons.Left)
            {
                _panning = false;
                Cursor = Cursors.Default;
            }
        }

        protected override void OnMouseDoubleClick(MouseEventArgs e)
        {
            base.OnMouseDoubleClick(e);
            if (e.Button != MouseButtons.Left) return;
            if (Math.Abs(_zoom - 1f) < 0.001f) Fit();
            else ActualSize();
        }

        protected override bool IsInputKey(Keys keyData)
        {
            switch (keyData)
            {
                case Keys.Left:
                case Keys.Right:
                case Keys.Up:
                case Keys.Down:
                    return true;
            }
            return base.IsInputKey(keyData);
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            base.OnKeyDown(e);
            int step = 80;
            switch (e.KeyCode)
            {
                case Keys.Escape:
                    Form owner = FindForm();
                    if (owner != null) owner.Close();
                    e.Handled = true;
                    break;
                case Keys.D0:
                case Keys.NumPad0:
                    Fit();
                    e.Handled = true;
                    break;
                case Keys.D1:
                case Keys.NumPad1:
                    ActualSize();
                    e.Handled = true;
                    break;
                case Keys.Oemplus:
                case Keys.Add:
                    ZoomBy(1.25f);
                    e.Handled = true;
                    break;
                case Keys.OemMinus:
                case Keys.Subtract:
                    ZoomBy(0.8f);
                    e.Handled = true;
                    break;
                case Keys.Left:
                    _ox += step; Invalidate(); e.Handled = true; break;
                case Keys.Right:
                    _ox -= step; Invalidate(); e.Handled = true; break;
                case Keys.Up:
                    _oy += step; Invalidate(); e.Handled = true; break;
                case Keys.Down:
                    _oy -= step; Invalidate(); e.Handled = true; break;
                case Keys.T:
                    if (e.Control)
                    {
                        Form f = FindForm();
                        if (f != null) f.TopMost = !f.TopMost;
                        e.Handled = true;
                    }
                    break;
            }
        }
    }

    #endregion

    #region 预览窗口

    internal sealed class PreviewForm : Form
    {
        private readonly Image _img;
        private readonly ZoomCanvas _canvas;
        private readonly string _baseTitle;

        public PreviewForm(Image img, string title)
        {
            _img = img;
            _baseTitle = title;

            Text = title;
            StartPosition = FormStartPosition.CenterScreen;
            TopMost = true;
            KeyPreview = true;
            MinimumSize = new Size(380, 280);
            DoubleBuffered = true;

            Rectangle wa = Screen.PrimaryScreen.WorkingArea;
            int w = Math.Min(1040, wa.Width - 120);
            int h = Math.Min(760, wa.Height - 120);
            Size = new Size(Math.Max(w, 380), Math.Max(h, 280));

            _canvas = new ZoomCanvas();
            _canvas.Dock = DockStyle.Fill;
            _canvas.ZoomChanged += OnZoomChanged;
            Controls.Add(_canvas);
            _canvas.SetImage(img);
        }

        private void OnZoomChanged(object sender, EventArgs e)
        {
            try
            {
                Text = string.Format("{0}    [{1:0}%]", _baseTitle, _canvas.Zoom * 100f);
            }
            catch { }
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            _canvas.InitialView();
            _canvas.Focus();
            OnZoomChanged(this, EventArgs.Empty);
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            base.OnFormClosed(e);
            try { _img.Dispose(); } catch { }
        }
    }

    #endregion

    #region 加载项主体

    // ★ ClassInterfaceType 必须是 None。
    //
    // AutoDual 会为类自动生成一个 `_Connect` 类接口，而它的 vtable 把
    // System.Object 的公共方法排在前面（实测 tools/typedump.py）：
    //     slot 7 ToString / slot 8 Equals / slot 9 GetHashCode / slot 10 GetType
    //     slot 11 OnConnection / 12 OnDisconnection / ...
    // Office 按 dual 约定调 slot 7 期望拿到 OnConnection，实际调到 ToString，
    // 签名对不上 -> CLR 致命错误 0x80131506 -> 宿主进程崩溃
    // （dotnet/runtime#38950 记录的正是 "instead of calling OnConnection(),
    //   it is calling the object's ToString()"）。
    //
    // None 表示不生成任何类接口：QI(IDispatch) 与 QI(IDTExtensibility2)
    // 都会落到我们那个 dual 接口上，slot 7 == OnConnection，两种调用方式都安全。
    // 已用 tools/invoketest.py vtable 实测：slot 7 能完整跑通 OnConnection。
    [ComVisible(true)]
    [Guid("4F2B8A61-9C3D-4C1E-9F1A-7D2E5B6A1C08")]
    [ProgId("ImageZoomAddin.Connect")]
    [ClassInterface(ClassInterfaceType.None)]
    public class Connect : IDTExtensibility2
    {
        private object _appRaw;
        private dynamic _app;
        private string _host = "";
        private Timer _poll;
        private MagnifierButton _magnifier;   // 放大镜（预览原图）
        private MagnifierButton _crop;        // 裁剪（Word / Excel 都有）
        private MagnifierButton _rotate;      // 旋转（Word / Excel 都有）
        private CropMenuForm _menu;           // 裁剪的二级菜单（按形状 / 按比例 / 重置）
        private RotateMenuForm _rotateMenu;   // 旋转的二级菜单（左右 90° / 水平垂直翻转）
        private PreviewForm _preview;
        private string _currentKey = "";
        private Rectangle _lastRect = Rectangle.Empty;   // 上次定位时图片在屏幕上的矩形
        private int _tickBusy;
        private int _diagTicks;

        // 探针：CoCreateInstance 一成功就会执行到这里。
        // 日志里有 ctor 但没有 OnConnection => 要么 vtable 布局不对（坑 3），
        // 要么 Office 没认下 IDTExtensibility2。vtable 用 tools/vtabletest.py 单独验。
        public Connect()
        {
            Log.Write("ctor: add-in object created");
        }

        // ------------------------------------------------------------ 生命周期

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            // 第一行就写日志：只要这行出现，就说明 Office 真的调到了 onConnection
            // （vtable 布局正确），后面任何一步出问题都能从日志定位。
            Log.Write("OnConnection: enter, mode=" + Convert.ToString(ConnectMode));
            try
            {
                // 宿主白名单：不是 Word / Excel 就什么都不做。
                // WPS 自带图片放大预览，且可能复用 Office 的 Addins 注册把本加载项拉起来，
                // 这里直接返回 —— 不建 Timer、不建窗口、不轮询，等于完全不介入。
                string host;
                if (!IsSupportedHost(out host))
                {
                    Log.Write("OnConnection: 跳过（宿主不受支持：" + CurrentProcessName() + "）");
                    return;
                }

                _appRaw = Application;
                Log.Write("OnConnection: 1 app 指针已收到");

                _app = Application;
                Log.Write("OnConnection: 2 dynamic 绑定完成");

                _host = host;
                Log.Write("OnConnection: 3 host=" + _host);

                StartPolling();
                Log.Write("OnConnection: done");
            }
            catch (Exception ex)
            {
                Log.Write("OnConnection 失败", ex);
            }
        }

        /// <summary>幂等：重复调用只启动一次。所有异常都吞掉，绝不让它冒到 Office 里。</summary>
        private void StartPolling()
        {
            try
            {
                if (_poll != null) return;

                // 宿主不是 Word/Excel（例如 WPS）时 _host 为空，这里什么都不启动。
                // OnStartupComplete 会无条件调一次本方法，所以这道保险必须有。
                if (string.IsNullOrEmpty(_host))
                {
                    Log.Write("StartPolling: 跳过（宿主不受支持）");
                    return;
                }

                Log.Write("StartPolling: creating timer");
                Timer t = new Timer();
                t.Interval = 200;
                t.Tick += OnPollTick;
                _poll = t;
                Log.Write("StartPolling: timer created");

                t.Start();
                Log.Write("StartPolling: timer started");
            }
            catch (Exception ex)
            {
                Log.Write("StartPolling 失败", ex);
            }
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            try
            {
                Log.Write("OnDisconnection: " + RemoveMode);
                if (_poll != null)
                {
                    _poll.Stop();
                    _poll.Tick -= OnPollTick;
                    _poll.Dispose();
                    _poll = null;
                }
                HideMagnifier();
                ClosePreview();
                _app = null;
                if (_appRaw != null)
                {
                    try { Marshal.ReleaseComObject(_appRaw); } catch { }
                    _appRaw = null;
                }
            }
            catch (Exception ex) { Log.Write("OnDisconnection 失败", ex); }
        }

        public void OnAddInsUpdate(ref Array custom) { }

        public void OnStartupComplete(ref Array custom)
        {
            Log.Write("OnStartupComplete");
            // 兜底：万一 OnConnection 阶段轮询没起来（异常/宿主状态），这里再试一次。
            StartPolling();
        }

        public void OnBeginShutdown(ref Array custom) { Log.Write("OnBeginShutdown"); }

        /// <summary>当前宿主进程名（不含扩展名），如 WINWORD / EXCEL / wps / et / wpp。</summary>
        private static string CurrentProcessName()
        {
            try { return System.Diagnostics.Process.GetCurrentProcess().ProcessName ?? ""; }
            catch { return ""; }
        }

        /// <summary>
        /// 只服务真正的 Microsoft Office 桌面宿主（WINWORD.EXE / EXCEL.EXE / POWERPNT.EXE）。
        ///
        /// 为什么不靠 Application.Name 判断：WPS 等第三方 Office 兼容层会去读
        /// HKCU\Software\Microsoft\Office\*\Addins 这份注册表，从而把本加载项也拉起来；
        /// 它们暴露的 Application.Name 五花八门（"WPS Writer"/"WPS 表格"…），
        /// 用名字匹配容易漏。直接看进程名最干脆 —— WPS 自带图片放大预览，
        /// 不需要我们介入，这里一刀切掉。
        /// </summary>
        private static bool IsSupportedHost(out string host)
        {
            host = "";
            string name = CurrentProcessName();
            if (name.Equals("WINWORD", StringComparison.OrdinalIgnoreCase)) { host = "Word"; return true; }
            if (name.Equals("EXCEL", StringComparison.OrdinalIgnoreCase)) { host = "Excel"; return true; }
            if (name.Equals("POWERPNT", StringComparison.OrdinalIgnoreCase)) { host = "PowerPoint"; return true; }
            return false;
        }

        // ------------------------------------------------------------ 轮询选中

        private void OnPollTick(object sender, EventArgs e)
        {
            if (System.Threading.Interlocked.Exchange(ref _tickBusy, 1) == 1) return;
            try
            {
                if (_preview != null && !_preview.IsDisposed && _preview.Visible)
                {
                    HideMagnifier();
                    return;
                }
                if (!Native.IsOwnProcessForeground())
                {
                    HideMagnifier();
                    return;
                }

                // 前 5 次轮询打印选择状态，便于首次部署时定位"检测不到图片"的问题
                if (_diagTicks < 5)
                {
                    _diagTicks++;
                    Log.Write("diag#" + _diagTicks + " " + DescribeSelection());
                }

                PictureHit hit = DetectPicture();
                if (hit == null)
                {
                    HideMagnifier();
                    return;
                }
                ShowMagnifier(hit);
            }
            catch (Exception ex)
            {
                Log.Write("轮询异常", ex);
                HideMagnifier();
            }
            finally
            {
                System.Threading.Interlocked.Exchange(ref _tickBusy, 0);
            }
        }

        private string DescribeSelection()
        {
            try
            {
                // PowerPoint 的 Application 上没有 Selection（实测 MissingMemberException），
                // 选中状态挂在 DocumentWindow 上，所以这里必须先拿 ActiveWindow。
                if (_host == "PowerPoint")
                {
                    dynamic win = null;
                    try { win = _app.ActiveWindow; } catch { return "PowerPoint.ActiveWindow=null(可能是放映中)"; }
                    if (win == null) return "PowerPoint.ActiveWindow=null";
                    dynamic sel = win.Selection;
                    string s = "PowerPoint.Selection.Type=" + Convert.ToString(sel.Type) +
                               " ViewType=" + Convert.ToString(win.ViewType);
                    try { s += " ShapeRange.Count=" + Convert.ToString(sel.ShapeRange.Count); }
                    catch { s += " ShapeRange=n/a"; }
                    return s;
                }

                dynamic sel2 = _app.Selection;
                if (sel2 == null) return "selection=null";

                string t;

                // Excel 的 Selection 没有 Type 属性（只有 Word 有），单独处理
                if (_host == "Excel")
                {
                    t = "Excel.Selection";
                    try { t += " ShapeRange.Count=" + Convert.ToString(sel2.ShapeRange.Count); }
                    catch { t += " ShapeRange=n/a(选中的是单元格)"; }
                    return t;
                }

                t = "Selection.Type=" + Convert.ToString(sel2.Type);

                if (_host == "Word")
                {
                    try { t += " InlineShapes=" + Convert.ToString(sel2.Range.InlineShapes.Count); }
                    catch { }
                }
                return t;
            }
            catch (Exception ex) { return "DescribeSelection failed: " + ex.Message; }
        }

        private PictureHit DetectPicture()
        {
            try
            {
                if (_app == null) return null;
                if (_host == "Word") return DetectWordPicture();
                if (_host == "Excel") return DetectExcelPicture();
                if (_host == "PowerPoint") return DetectPowerPointPicture();
            }
            catch (Exception ex)
            {
                Log.Write("检测选中图片失败", ex);
            }
            return null;
        }

        private PictureHit DetectWordPicture()
        {
            dynamic sel = _app.Selection;
            if (sel == null) return null;

            int type = Convert.ToInt32(sel.Type);

            // wdSelectionInlineShape = 7
            // ★ 这里曾误写成 3（3 其实是 wdSelectionFrame），导致选中图片后永远匹配不上。
            //   Word 实测：选中嵌入式图片时 Selection.Type 返回 7。
            // ★ 另一个坑：Selection 对象没有 InlineShape（单数）属性，
            //   早期写的 sel.InlineShape 会抛 RuntimeBinderException，
            //   内嵌图形只能从 InlineShapes 集合里取。
            if (type == 7)
            {
                try
                {
                    dynamic inlines = sel.InlineShapes;
                    if (inlines == null) return null;
                    if (Convert.ToInt32(inlines.Count) < 1) return null;

                    dynamic shape = inlines.Item(1);

                    PictureHit hit = new PictureHit();
                    hit.Host = "Word";
                    hit.IsInline = true;
                    hit.Inline = shape;
                    hit.Name = "嵌入式图片";
                    hit.Key = "W:I:" + Convert.ToString(sel.Range.Start);
                    hit.ScreenRect = GetWordRangeRect(sel.Range);
                    return hit;
                }
                catch (Exception ex)
                {
                    Log.Write("取内嵌图片失败", ex);
                    return null;
                }
            }

            // wdSelectionShape = 8（浮动图片）
            if (type == 8)
            {
                dynamic range = sel.ShapeRange;
                if (range == null) return null;
                if (Convert.ToInt32(range.Count) < 1) return null;

                dynamic shape = range.Item(1);
                int shapeType = Convert.ToInt32(shape.Type);
                // msoLinkedPicture = 11, msoPicture = 13
                if (shapeType != 11 && shapeType != 13) return null;

                PictureHit hit = new PictureHit();
                hit.Host = "Word";
                hit.IsInline = false;
                hit.Shape = shape;
                hit.Name = Convert.ToString(shape.Name);
                hit.Key = "W:S:" + hit.Name + ":" + Convert.ToString(shape.Anchor.Start);
                hit.ScreenRect = GetWordRangeRect(shape.Anchor);
                return hit;
            }

            return null;
        }

        private PictureHit DetectExcelPicture()
        {
            dynamic sel = _app.Selection;
            if (sel == null) return null;

            dynamic range = null;
            try { range = sel.ShapeRange; }
            catch { return null; }

            if (range == null) return null;
            if (Convert.ToInt32(range.Count) < 1) return null;

            dynamic shape = range.Item(1);
            int shapeType = Convert.ToInt32(shape.Type);
            // 11 = msoLinkedPicture, 13 = msoPicture
            bool isPicture = (shapeType == 11 || shapeType == 13);
            // 本加载项「裁剪为形状」生成的图片填充自选图形（Excel 没法把图片真的变成形状，
            // 只能新建一个自选图形再把图片填进去），也要继续认作"图片"，
            // 否则裁完形状放大镜和裁剪按钮就一起消失了。
            bool isShapeImage = !isPicture && IsShapeImage(shape);
            if (!isPicture && !isShapeImage) return null;

            PictureHit hit = new PictureHit();
            hit.Host = "Excel";
            hit.IsShapeImage = isShapeImage;
            hit.Shape = shape;
            hit.Name = Convert.ToString(shape.Name);
            try { hit.SheetName = Convert.ToString(_app.ActiveSheet.Name); }
            catch { hit.SheetName = ""; }
            hit.Key = "E:" + hit.SheetName + ":" + hit.Name;
            hit.ScreenRect = GetExcelShapeRect(shape);
            return hit;
        }

        /// <summary>
        /// PowerPoint 的选中图片。
        ///
        /// 两个和 Word/Excel 都不一样的地方：
        ///   1) PowerPoint **没有 Application.Selection**（实测 MissingMemberException），
        ///      选中对象挂在 DocumentWindow 上：ActiveWindow.Selection。
        ///   2) "选中形状"的 Selection.Type 是 2（ppSelectionShapes），不是 Word 那种 7/8；
        ///      放映时 ActiveWindow 直接不可用，这里必须全程容错。
        /// </summary>
        private PictureHit DetectPowerPointPicture()
        {
            dynamic win = null;
            try { win = _app.ActiveWindow; }
            catch { return null; }                 // 放映中 / 没有打开的窗口
            if (win == null) return null;

            // 只在"普通视图 / 幻灯片视图"下工作；浏览视图、备注视图里没有可操作的对象
            try
            {
                int vt = Convert.ToInt32(win.ViewType);
                if (vt != 9 && vt != 1) return null;
            }
            catch { }

            dynamic sel = null;
            try { sel = win.Selection; }
            catch { return null; }
            if (sel == null) return null;

            // ppSelectionShapes = 2；0=无 1=幻灯片 3=文本
            try { if (Convert.ToInt32(sel.Type) != 2) return null; }
            catch { return null; }

            dynamic range = null;
            try { range = sel.ShapeRange; }
            catch { return null; }
            if (range == null) return null;
            if (Convert.ToInt32(range.Count) < 1) return null;

            dynamic shape = range.Item(1);
            int shapeType = Convert.ToInt32(shape.Type);

            // 11 = msoLinkedPicture, 13 = msoPicture
            // 14 = msoPlaceholder：内容占位符里插了图片时，ContainedType 才是 13
            bool isPicture = (shapeType == 11 || shapeType == 13);
            if (!isPicture && shapeType == 14)
            {
                try { isPicture = Convert.ToInt32(shape.PlaceholderFormat.ContainedType) == 13; }
                catch { }
            }
            if (!isPicture) return null;

            // ★ 切页后 Selection 仍可能指着上一页那张图（实测：GotoSlide 之后
            //   Selection.Type 还是 2、ShapeRange 还是旧的 shape），
            //   不核对的话按钮会赖在已经翻走的那一页上。
            try
            {
                int viewSlideId = Convert.ToInt32(win.View.Slide.SlideID);
                int shapeSlideId = Convert.ToInt32(shape.Parent.SlideID);
                if (viewSlideId != shapeSlideId) return null;
            }
            catch { }

            PictureHit hit = new PictureHit();
            hit.Host = "PowerPoint";
            hit.Shape = shape;
            hit.Name = Convert.ToString(shape.Name);

            // Key 用「幻灯片 ID + 形状 Id」：名字可以重名、也可以被用户改掉，
            // 而 Shape.Id 直接对应 OOXML 里的 cNvPr/@id，稳定且能用来定位包内的图片。
            int shapeId = 0, slideId = 0;
            try { shapeId = Convert.ToInt32(shape.Id); } catch { }
            try { slideId = Convert.ToInt32(win.View.Slide.SlideID); } catch { }
            hit.ShapeId = shapeId;
            try { hit.SlideIndex = Convert.ToInt32(win.View.Slide.SlideIndex); } catch { hit.SlideIndex = 0; }
            hit.Key = "P:" + slideId + ":" + shapeId + ":" + hit.Name;
            hit.ScreenRect = GetPowerPointShapeRect(shape);
            return hit;
        }

        // ------------------------------------------------------------ 屏幕坐标

        private Rectangle GetWordRangeRect(dynamic range)
        {
            try
            {
                dynamic win = _app.ActiveWindow;
                if (win == null) return Rectangle.Empty;

                dynamic left = 0;
                dynamic top = 0;
                dynamic width = 0;
                dynamic height = 0;

                win.GetPoint(out left, out top, out width, out height, range);

                int l = Convert.ToInt32(left);
                int t = Convert.ToInt32(top);
                int w = Convert.ToInt32(width);
                int h = Convert.ToInt32(height);

                if (w <= 0 || h <= 0) return Rectangle.Empty;
                if (l < -20000 || t < -20000 || l > 30000 || t > 30000) return Rectangle.Empty;
                return new Rectangle(l, t, w, h);
            }
            catch (Exception ex)
            {
                Log.Write("Word GetPoint 失败（改用鼠标位置定位）", ex);
                return Rectangle.Empty;
            }
        }

        /// <summary>
        /// Excel 里那张图片在屏幕上的矩形。
        ///
        /// ★★ 天坑：Excel 的 Window.PointsToScreenPixelsX/Y **不按屏幕 DPI 换算**。
        ///   实测（tools/xlrect.cs，用纯品红图片 + 截屏找色块做基准）：
        ///     f(p) 恒等于 p + f(0)，斜率 1.000 —— 100% 缩放下 600 磅的图片算出来 600 像素，
        ///     而它实际占 800 像素；50% 缩放时真实斜率 0.667、150% 时 2.000，f() 仍是 1.000。
        ///   也就是说自己必须乘上 (屏幕DPI / 72) × (Zoom / 100)：
        ///     96 DPI、100% → 1.3333；50% → 0.6667；150% → 2.0000（都与实测吻合）
        ///   f(0) 本身倒是对的：它就是工作表原点（A1 左上角）在屏幕上的像素坐标，
        ///   连滚动和缩放带来的偏移都算准了，可以直接当锚点用。
        ///
        ///   不校正的后果就是用户看到的样子 —— 按钮只走了 3/4 的距离，
        ///   落在图片偏左上、也就是"放大镜跑到图片里面"。
        /// </summary>
        private Rectangle GetExcelShapeRect(dynamic shape)
        {
            try
            {
                dynamic win = _app.ActiveWindow;
                if (win == null) return Rectangle.Empty;

                double left = Convert.ToDouble(shape.Left);
                double top = Convert.ToDouble(shape.Top);
                double w = Convert.ToDouble(shape.Width);
                double h = Convert.ToDouble(shape.Height);

                // 图片被旋转过（Shape.Rotation ≠ 0）时，Left/Top/Width/Height 还是"没转之前"的框，
                // 屏幕上看到的却是绕中心转过的外接矩形 —— 不折算按钮就会贴着旧框跑。
                double rot = 0;
                try { rot = Convert.ToDouble(shape.Rotation); } catch { }
                if (Math.Abs(rot) > 0.5)
                {
                    double rad = rot * Math.PI / 180.0;
                    double c = Math.Abs(Math.Cos(rad)), s = Math.Abs(Math.Sin(rad));
                    double bw = w * c + h * s;
                    double bh = w * s + h * c;
                    double cx = left + w / 2.0, cy = top + h / 2.0;
                    left = cx - bw / 2.0;
                    top = cy - bh / 2.0;
                    w = bw;
                    h = bh;
                }

                int ox = Convert.ToInt32(win.PointsToScreenPixelsX(0));
                int oy = Convert.ToInt32(win.PointsToScreenPixelsY(0));

                // 窗口还没真正画出来时 f() 一律返回 0（启动后的头一两秒会这样），
                // 这时算出来的是垃圾坐标，宁可不显示按钮，等下一轮轮询。
                int probeX = Convert.ToInt32(win.PointsToScreenPixelsX(100));
                int probeY = Convert.ToInt32(win.PointsToScreenPixelsY(100));
                if (ox == 0 && oy == 0 && probeX == 0 && probeY == 0) return Rectangle.Empty;

                int dpiX, dpiY;
                Native.GetDpi(out dpiX, out dpiY);
                double zoom = ExcelZoom(win);
                double kx = dpiX / 72.0 * zoom;
                double ky = dpiY / 72.0 * zoom;

                int x1 = ox + (int)Math.Round(left * kx);
                int y1 = oy + (int)Math.Round(top * ky);
                int rw = (int)Math.Round(w * kx);
                int rh = (int)Math.Round(h * ky);

                if (rw <= 0 || rh <= 0) return Rectangle.Empty;
                if (x1 < -20000 || y1 < -20000 || x1 > 30000 || y1 > 30000) return Rectangle.Empty;
                return new Rectangle(x1, y1, rw, rh);
            }
            catch (Exception ex)
            {
                Log.Write("Excel 坐标换算失败（改用鼠标位置定位）", ex);
                return Rectangle.Empty;
            }
        }

        /// <summary>工作表缩放比例（1.0 = 100%）。</summary>
        private static double ExcelZoom(dynamic win)
        {
            try
            {
                double z = Convert.ToDouble(win.Zoom);
                if (z > 10 && z < 400) return z / 100.0;
            }
            catch { }
            return 1.0;
        }

        /// <summary>
        /// PowerPoint 里那张图片在屏幕上的矩形。
        ///
        /// ★★★ 这里和 Excel 是**相反**的，绝不能照抄 Excel 的写法：
        ///   Excel 的 PointsToScreenPixelsX/Y 斜率恒为 1.000（只做磅→像素的 1:1 映射，
        ///   既不认 DPI 也不认工作表缩放），所以 Excel 那边必须自己乘 dpi/72 × zoom。
        ///   PowerPoint 的斜率**直接就是 dpi/72 × 视图缩放** —— 实测（tools/pptprobe.cs：
        ///   拿 DocumentWindow.RangeFromPoint 从图片中心向四边逐像素外扩，反量真实矩形）：
        ///     600 磅的图，50% 缩放 → 400px、150% → 1200px、200% → 1600px，
        ///     而 f(L+W)-f(L) 每次都跟真实值一致。
        ///   所以这边**直接用 f(Left)、f(Left+Width)** 相减就行，一个系数都不要乘。
        ///
        /// 这也顺带治好了「适应窗口」的坑：PowerPoint 新建演示文稿默认按窗口自适应，
        /// View.Zoom 只报整数（105%），内部实际可能是 104.1%，自己乘 1.4 会算出 840px
        /// 而实际只有 833px —— 按钮就偏出去 7 个像素。交给 f() 算就没有这个偏差。
        /// </summary>
        private Rectangle GetPowerPointShapeRect(dynamic shape)
        {
            try
            {
                dynamic win = null;
                try { win = _app.ActiveWindow; } catch { return Rectangle.Empty; }
                if (win == null) return Rectangle.Empty;

                double left = Convert.ToDouble(shape.Left);
                double top = Convert.ToDouble(shape.Top);
                double w = Convert.ToDouble(shape.Width);
                double h = Convert.ToDouble(shape.Height);

                // 旋转后 Left/Top/Width/Height 仍是"没转之前"的框，屏幕上看到的是
                // 绕中心转过的外接矩形 —— 不折算按钮就会贴着旧框跑（和 Excel 同样的处理）。
                double rot = 0;
                try { rot = Convert.ToDouble(shape.Rotation); } catch { }
                if (Math.Abs(rot) > 0.5)
                {
                    double rad = rot * Math.PI / 180.0;
                    double c = Math.Abs(Math.Cos(rad)), s = Math.Abs(Math.Sin(rad));
                    double bw = w * c + h * s;
                    double bh = w * s + h * c;
                    double cx = left + w / 2.0, cy = top + h / 2.0;
                    left = cx - bw / 2.0;
                    top = cy - bh / 2.0;
                    w = bw;
                    h = bh;
                }

                // 窗口还没真正画出来时 f() 一律返回 0（启动后的头一两秒会这样），
                // 这时算出来的是垃圾坐标，宁可不显示按钮，等下一轮轮询。
                int ox = Convert.ToInt32(win.PointsToScreenPixelsX(0f));
                int oy = Convert.ToInt32(win.PointsToScreenPixelsY(0f));
                int probeX = Convert.ToInt32(win.PointsToScreenPixelsX(100f));
                int probeY = Convert.ToInt32(win.PointsToScreenPixelsY(100f));
                if (ox == 0 && oy == 0 && probeX == 0 && probeY == 0) return Rectangle.Empty;

                int x1 = Convert.ToInt32(win.PointsToScreenPixelsX((float)left));
                int y1 = Convert.ToInt32(win.PointsToScreenPixelsY((float)top));
                int x2 = Convert.ToInt32(win.PointsToScreenPixelsX((float)(left + w)));
                int y2 = Convert.ToInt32(win.PointsToScreenPixelsY((float)(top + h)));

                int rw = x2 - x1;
                int rh = y2 - y1;

                if (rw <= 0 || rh <= 0) return Rectangle.Empty;
                if (x1 < -20000 || y1 < -20000 || x1 > 30000 || y1 > 30000) return Rectangle.Empty;
                return new Rectangle(x1, y1, rw, rh);
            }
            catch (Exception ex)
            {
                Log.Write("PowerPoint 坐标换算失败（改用鼠标位置定位）", ex);
                return Rectangle.Empty;
            }
        }

        // ------------------------------------------------------------ 浮动按钮显示

        private const int BtnSize = MagnifierButton.BoxSize;   // 圆角方框边长
        private const int BtnGap = 7;      // 按钮之间的竖向间距
        private const int BtnInset = 8;    // 距图片边缘的内缩

        private void ShowMagnifier(PictureHit hit)
        {
            Rectangle target = hit.ScreenRect;
            if (target.IsEmpty)
            {
                Point cursor = Cursor.Position;
                target = new Rectangle(cursor.X, cursor.Y, 1, 1);
            }

            if (_magnifier == null || _magnifier.IsDisposed)
            {
                _magnifier = new MagnifierButton();
                _magnifier.Clicked += OnMagnifierClicked;
            }
            if (_crop == null || _crop.IsDisposed)
            {
                _crop = new MagnifierButton(ButtonKind.Crop);
                _crop.Clicked += OnCropClicked;
            }
            if (_rotate == null || _rotate.IsDisposed)
            {
                _rotate = new MagnifierButton(ButtonKind.Rotate);
                _rotate.Clicked += OnRotateClicked;
            }

            // 四个条件都满足才跳过重排：同一张图 + 三个按钮都在 + 图片在屏幕上的矩形没变。
            // ★ 必须比较 ScreenRect：以前只比 Key，导致"缩放/拖动/滚动图片后按钮不动"
            //   （Excel 里缩放图片时选中对象没变，Key 自然也不变）。
            bool allVisible = _magnifier.Visible
                              && _crop != null && !_crop.IsDisposed && _crop.Visible
                              && _rotate != null && !_rotate.IsDisposed && _rotate.Visible;
            bool sameTarget = allVisible
                              && string.Equals(_currentKey, hit.Key, StringComparison.Ordinal)
                              && SameRect(_lastRect, target);
            if (sameTarget) return;

            Point[] pos = PlaceButtons(target, 3);

            bool magWasVisible = _magnifier.Visible;
            _magnifier.Location = pos[0];
            if (!magWasVisible) _magnifier.Show();
            _magnifier.Repaint();

            _crop.Location = pos[1];
            if (!_crop.Visible) _crop.Show();
            _crop.Repaint();

            _rotate.Location = pos[2];
            if (!_rotate.Visible) _rotate.Show();
            _rotate.Repaint();

            _currentKey = hit.Key;
            _lastRect = target;
            Log.Write("magnifier: " + hit.Key + " rect=" + target + " -> mag=" + pos[0] + " crop=" + pos[1] +
                      " rot=" + pos[2] + (hit.IsShapeImage ? " (形状图)" : ""));
        }

        /// <summary>两个矩形是不是"同一个位置"（差 1px 以内算没动，避免像素取整导致每轮都重排）。</summary>
        private static bool SameRect(Rectangle a, Rectangle b)
        {
            if (a.IsEmpty || b.IsEmpty) return false;
            return Math.Abs(a.X - b.X) <= 1 && Math.Abs(a.Y - b.Y) <= 1
                   && Math.Abs(a.Width - b.Width) <= 1 && Math.Abs(a.Height - b.Height) <= 1;
        }

        /// <summary>
        /// 竖排摆在图片右下角：放大镜在最上面（贴住右下角），裁剪紧接在它下方。
        /// 优先放在图片右侧外面 —— 图片内侧的右下角在裁剪状态下正好是拖动用的手柄，
        /// 按钮压上去会挡住拖拽；实在没地方（图片贴着屏幕右边）才缩回图片内侧。
        /// 返回每个按钮的左上角，顺序与调用方一致。
        /// </summary>
        private static Point[] PlaceButtons(Rectangle target, int count)
        {
            if (target.Width < 1) target.Width = 1;
            if (target.Height < 1) target.Height = 1;

            Rectangle area = Screen.FromRectangle(target).WorkingArea;
            if (area.IsEmpty) area = Screen.PrimaryScreen.WorkingArea;

            // 图片可能有一部分滚出屏幕：按可见部分定位，否则按钮会跑到屏幕外
            Rectangle vis = Rectangle.Intersect(target, area);
            Rectangle anchor = (vis.Width >= 8 && vis.Height >= 8) ? vis : target;

            int groupH = count * BtnSize + (count - 1) * BtnGap;

            int x = anchor.Right + BtnGap;                 // 图片右侧外面
            if (x + BtnSize > area.Right - 2)              // 右边放不下 → 缩回图片内侧右下角
                x = anchor.Right - BtnInset - BtnSize;

            int y = anchor.Bottom - BtnInset - BtnSize;     // 放大镜贴右下角

            if (y + groupH > area.Bottom - 2) y = area.Bottom - groupH - 2;
            if (y < area.Top + 2) y = area.Top + 2;
            if (y + BtnSize > area.Bottom - 2) y = area.Bottom - BtnSize - 2;   // 屏幕太矮

            if (x < area.Left + 2) x = area.Left + 2;
            if (x + BtnSize > area.Right - 2) x = area.Right - BtnSize - 2;

            Point[] res = new Point[count];
            for (int i = 0; i < count; i++) res[i] = new Point(x, y + i * (BtnSize + BtnGap));
            return res;
        }

        private void HideMagnifier()
        {
            _currentKey = "";
            _lastRect = Rectangle.Empty;
            CloseCropMenu();
            CloseRotateMenu();
            try
            {
                if (_magnifier != null && !_magnifier.IsDisposed && _magnifier.Visible)
                    _magnifier.Hide();
            }
            catch { }
            try
            {
                if (_crop != null && !_crop.IsDisposed && _crop.Visible)
                    _crop.Hide();
            }
            catch { }
            try
            {
                if (_rotate != null && !_rotate.IsDisposed && _rotate.Visible)
                    _rotate.Hide();
            }
            catch { }
        }

        private void OnMagnifierClicked(object sender, EventArgs e)
        {
            try
            {
                CloseCropMenu();   // 预览窗打开时把裁剪菜单收掉，避免两个浮层叠在一起
                CloseRotateMenu();

                PictureHit hit = DetectPicture();
                if (hit == null)
                {
                    HideMagnifier();
                    return;
                }

                Cursor previous = Cursor.Current;
                Cursor.Current = Cursors.WaitCursor;
                try
                {
                    ImagePayload payload = Extract(hit);
                    if (payload == null || payload.Bytes == null || payload.Bytes.Length == 0)
                    {
                        MessageBox.Show(
                            "没能取到这张图片的原始数据。\n\n" +
                            "可能原因：图片是链接到外部文件的、是 OLE 对象而不是普通图片，或文档尚未保存过。",
                            "图片放大预览", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    Bitmap bmp = Decode(payload.Bytes);
                    if (bmp == null)
                    {
                        MessageBox.Show("图片数据解码失败。", "图片放大预览", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    string display = "";
                    if (!hit.ScreenRect.IsEmpty)
                        display = string.Format(" · 文档中显示 {0}×{1}px", hit.ScreenRect.Width, hit.ScreenRect.Height);

                    string title = string.Format("{0}  原图 {1}×{2}px{3} · 来源：{4}",
                        string.IsNullOrEmpty(hit.Name) ? "图片" : hit.Name,
                        bmp.Width, bmp.Height, display, payload.Source);

                    ClosePreview();
                    _preview = new PreviewForm(bmp, title);
                    _preview.Show();
                    _preview.Activate();
                }
                finally
                {
                    Cursor.Current = previous;
                }
            }
            catch (Exception ex)
            {
                Log.Write("打开预览失败", ex);
                MessageBox.Show("打开预览时出错：" + ex.Message, "图片放大预览", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                HideMagnifier();
            }
        }

        /// <summary>
        /// 点击裁剪按钮：开 / 关右侧的二级菜单（WPS 的样子）。
        /// 菜单里的「自由裁剪」才是原来的 Word 原生裁剪状态，
        /// 「按形状裁剪 / 按比例裁剪 / 重设形状和大小」都在菜单里完成。
        /// </summary>
        private void OnCropClicked(object sender, EventArgs e)
        {
            try
            {
                if (_menu != null && !_menu.IsDisposed && _menu.Visible)
                {
                    CloseCropMenu();
                    return;
                }

                CloseRotateMenu();   // 两个面板不叠着出现

                PictureHit hit = DetectPicture();
                if (hit == null) return;

                ShowCropMenu(hit);
            }
            catch (Exception ex) { Log.Write("crop menu: 打开失败", ex); }
        }

        // ------------------------------------------------------------ 裁剪二级菜单

        private void ShowCropMenu(PictureHit hit)
        {
            CloseCropMenu();

            CropMenuForm menu = new CropMenuForm();
            menu.Chosen += OnCropChosen;
            menu.Relayout();

            Rectangle btn = (_crop != null && !_crop.IsDisposed && _crop.Visible)
                          ? _crop.Bounds
                          : new Rectangle(Cursor.Position, new Size(BtnSize, BtnSize));
            Rectangle area = Screen.FromRectangle(btn).WorkingArea;

            // 默认贴在裁剪按钮右侧；右边放不下就翻到左侧
            int x = btn.Right + 6;
            if (x + menu.Width > area.Right - 2) x = btn.Left - 6 - menu.Width;
            if (x < area.Left + 2) x = area.Left + 2;

            int y = btn.Top - 4;
            if (y + menu.Height > area.Bottom - 2) y = area.Bottom - menu.Height - 2;
            if (y < area.Top + 2) y = area.Top + 2;

            menu.Location = new Point(x, y);
            menu.KeepOpenRect = btn;
            _menu = menu;
            menu.Show();
            Log.Write("crop menu: open " + menu.Bounds + " for " + hit.Key);
        }

        private void CloseCropMenu()
        {
            try
            {
                if (_menu != null && !_menu.IsDisposed) _menu.Close();
            }
            catch { }
            _menu = null;
        }

        // ------------------------------------------------------------ 旋转二级菜单

        /// <summary>
        /// 点击旋转按钮：开 / 关右侧的小面板。
        /// 内容和 Office「图片格式 → 排列 → 旋转」下拉完全一致。
        /// </summary>
        private void OnRotateClicked(object sender, EventArgs e)
        {
            try
            {
                if (_rotateMenu != null && !_rotateMenu.IsDisposed && _rotateMenu.Visible)
                {
                    CloseRotateMenu();
                    return;
                }

                PictureHit hit = DetectPicture();
                if (hit == null) return;

                ShowRotateMenu(hit);
            }
            catch (Exception ex) { Log.Write("rotate menu: 打开失败", ex); }
        }

        private void ShowRotateMenu(PictureHit hit)
        {
            CloseCropMenu();
            CloseRotateMenu();

            RotateMenuForm menu = new RotateMenuForm();
            menu.Chosen += OnRotateChosen;

            Rectangle btn = (_rotate != null && !_rotate.IsDisposed && _rotate.Visible)
                          ? _rotate.Bounds
                          : new Rectangle(Cursor.Position, new Size(BtnSize, BtnSize));
            Rectangle area = Screen.FromRectangle(btn).WorkingArea;

            int x = btn.Right + 6;
            if (x + menu.Width > area.Right - 2) x = btn.Left - 6 - menu.Width;
            if (x < area.Left + 2) x = area.Left + 2;

            int y = btn.Top - 4;
            if (y + menu.Height > area.Bottom - 2) y = area.Bottom - menu.Height - 2;
            if (y < area.Top + 2) y = area.Top + 2;

            menu.Location = new Point(x, y);
            menu.KeepOpenRect = btn;
            _rotateMenu = menu;
            menu.Show();
            Log.Write("rotate menu: open " + menu.Bounds + " for " + hit.Key);
        }

        private void CloseRotateMenu()
        {
            try
            {
                if (_rotateMenu != null && !_rotateMenu.IsDisposed) _rotateMenu.Close();
            }
            catch { }
            _rotateMenu = null;
        }

        private void OnRotateChosen(RotateKind kind)
        {
            CloseRotateMenu();
            try
            {
                PictureHit hit = DetectPicture();
                if (hit == null) { HideMagnifier(); return; }

                if (string.Equals(hit.Host, "Excel", StringComparison.Ordinal) ||
                    string.Equals(hit.Host, "PowerPoint", StringComparison.Ordinal))
                    ApplyShapeRotate(hit, kind);
                else
                    ApplyWordRotate(hit, kind);
            }
            catch (Exception ex) { Log.Write("rotate action 失败: " + kind, ex); }

            // 尺寸/位置变了，下一轮轮询重新定位按钮
            _currentKey = "";
            _lastRect = Rectangle.Empty;
        }

        /// <summary>
        /// Excel / PowerPoint 共用的旋转与翻转：Shape.Rotation（顺时针角度，0~360）
        /// 和 Shape.Flip（msoFlipHorizontal=0 / msoFlipVertical=1），
        /// 跟「图片格式 → 排列 → 旋转」走的是同一套对象模型属性。
        ///
        /// PowerPoint 有个特点：Flip 之后 Rotation 会被它自己归一化
        /// （实测 90° 的图 Flip(0) 之后 Rotation 读回来是 270，视觉上是同一个结果）。
        /// 所以每次都"读当前值再累加"，绝不缓存角度状态。
        /// </summary>
        private void ApplyShapeRotate(PictureHit hit, RotateKind kind)
        {
            dynamic shp = hit.Shape;
            if (shp == null) return;

            switch (kind)
            {
                case RotateKind.Right90:
                    shp.Rotation = (Convert.ToDouble(shp.Rotation) + 90.0) % 360.0;
                    break;
                case RotateKind.Left90:
                    shp.Rotation = (Convert.ToDouble(shp.Rotation) + 270.0) % 360.0;
                    break;
                case RotateKind.FlipH:
                    shp.Flip(0);   // msoFlipHorizontal
                    break;
                case RotateKind.FlipV:
                    shp.Flip(1);   // msoFlipVertical
                    break;
            }

            Log.Write("rotate: " + hit.Host + " " + kind + " -> rotation=" +
                      Convert.ToDouble(shp.Rotation).ToString("0.0") + ", key=" + hit.Key);
        }

        /// <summary>
        /// Word 的旋转 / 翻转：图片没有 Rotation API（InlineShape 没有、Shape 有但浮动图还要另想办法），
        /// 而 Word 自己存旋转的方式就是在 <c>pic:spPr/a:xfrm</c> 上写 rot / flipH / flipV ——
        /// 和「裁剪为形状」同一条路：取 Range.WordOpenXML，改属性，再 InsertXML 写回。
        ///
        /// 嵌入型图片改完仍然是嵌入型（框不变、图片围着中心转，和 Word 自带旋转后的表现一致）；
        /// 浮动图片先转嵌入型再改，最后还原成浮动并恢复环绕方式。
        /// </summary>
        private void ApplyWordRotate(PictureHit hit, RotateKind kind)
        {
            bool floating = !hit.IsInline;
            int wrap = -1;
            dynamic shape = null;

            try
            {
                if (floating)
                {
                    shape = hit.Shape;
                    try { wrap = Convert.ToInt32(shape.WrapFormat.Type); } catch { }
                    shape = shape.ConvertToInlineShape();
                }
                else
                {
                    shape = hit.Inline;
                }

                dynamic range = shape.Range;
                int start = Convert.ToInt32(range.Start);
                string xml = Convert.ToString(range.WordOpenXML);

                if (string.IsNullOrEmpty(xml))
                {
                    Log.Write("rotate: 取不到 XML, key=" + hit.Key);
                    MessageBox.Show("这张图片取不到可以修改的结构，没法旋转。\n\n" +
                                    "可以改用 Word 自带的「图片格式 → 排列 → 旋转」。",
                                    "旋转图片", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                string patched = PatchTransform(xml, kind);
                if (patched == null)
                {
                    Log.Write("rotate: 没有 a:xfrm, key=" + hit.Key);
                    MessageBox.Show("这张图片的结构不标准，改不了旋转角。", "旋转图片",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                range.InsertXML(patched);
                Log.Write("rotate: Word " + kind + " ok, key=" + hit.Key);
                AfterInsert(start, floating, wrap);
            }
            catch (Exception ex)
            {
                Log.Write("rotate: Word " + kind + " 失败", ex);
                MessageBox.Show("旋转失败：" + ex.Message + "\n\n可以改用 Word 自带的旋转。",
                                "旋转图片", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        /// <summary>
        /// 改 XML 里第一个 <c>a:xfrm</c> 的 rot / flipH / flipV。
        /// rot 的单位是 1/60000 度（正数 = 顺时针），一周 21600000。
        /// 改完不落 rot="0"、flipH="0"，保持 XML 干净（Word 自己也是这么写的）。
        /// </summary>
        private static string PatchTransform(string xml, RotateKind kind)
        {
            Match m = Regex.Match(xml, "<a:xfrm(?<attr>[^>]*)>");
            if (!m.Success) return null;

            string attrs = m.Groups["attr"].Value;
            bool selfClose = attrs.EndsWith("/");
            if (selfClose) attrs = attrs.Substring(0, attrs.Length - 1);

            int rot = AttrInt(attrs, "rot");
            bool fh = AttrFlag(attrs, "flipH");
            bool fv = AttrFlag(attrs, "flipV");

            switch (kind)
            {
                case RotateKind.Right90: rot = (rot + 5400000) % 21600000; break;
                case RotateKind.Left90: rot = (rot + 21600000 - 5400000) % 21600000; break;
                case RotateKind.FlipH: fh = !fh; break;
                case RotateKind.FlipV: fv = !fv; break;
            }

            string rest = StripAttr(StripAttr(StripAttr(attrs, "rot"), "flipH"), "flipV").Trim();

            StringBuilder sb = new StringBuilder();
            sb.Append("<a:xfrm");
            if (rest.Length > 0) sb.Append(' ').Append(rest);
            if (rot != 0) sb.Append(" rot=\"").Append(rot).Append('"');
            if (fh) sb.Append(" flipH=\"1\"");
            if (fv) sb.Append(" flipV=\"1\"");
            sb.Append(selfClose ? "/>" : ">");

            return xml.Substring(0, m.Index) + sb.ToString() + xml.Substring(m.Index + m.Length);
        }

        private static int AttrInt(string attrs, string name)
        {
            Match m = Regex.Match(attrs, "\\b" + name + "=\"(-?\\d+)\"");
            if (!m.Success) return 0;
            int v;
            return int.TryParse(m.Groups[1].Value, out v) ? v : 0;
        }

        private static bool AttrFlag(string attrs, string name)
        {
            Match m = Regex.Match(attrs, "\\b" + name + "=\"([^\"]*)\"");
            if (!m.Success) return false;
            string v = m.Groups[1].Value.Trim();
            return v == "1" || v.Equals("true", StringComparison.OrdinalIgnoreCase);
        }

        private static string StripAttr(string attrs, string name)
        {
            return Regex.Replace(attrs, "\\s*\\b" + name + "=\"[^\"]*\"", "");
        }

        private void OnCropChosen(CropChoice choice)
        {
            CloseCropMenu();
            if (choice == null) return;

            try
            {
                // 面板是 NOACTIVATE 的，选中状态还在；再取一次是为了拿到最新的对象
                PictureHit hit = DetectPicture();
                if (hit == null) { HideMagnifier(); return; }

                if (string.Equals(hit.Host, "Excel", StringComparison.Ordinal))
                {
                    RunExcelCrop(hit, choice);
                }
                else if (string.Equals(hit.Host, "PowerPoint", StringComparison.Ordinal))
                {
                    // PowerPoint 的图片本来就能直接改几何（AutoShapeType），
                    // 所以不用像 Excel 那样先还原成原图；比例裁剪、自由裁剪、重设
                    // 走的都是和 Word 同一套 PictureFormat / Ribbon 命令。
                    switch (choice.Kind)
                    {
                        case CropKind.Shape: ApplyPowerPointShapeCrop(hit, choice.Prst, choice.Label); break;
                        case CropKind.Ratio: ApplyRatioCrop(hit, choice.Rw, choice.Rh, choice.Label); break;
                        case CropKind.FreeCrop: ExecuteNativeCrop(hit); break;
                        case CropKind.Reset: ResetPictureCmd(hit); break;
                    }
                }
                else
                {
                    switch (choice.Kind)
                    {
                        case CropKind.Shape: ApplyShapeCrop(hit, choice.Prst, choice.Label); break;
                        case CropKind.Ratio: ApplyRatioCrop(hit, choice.Rw, choice.Rh, choice.Label); break;
                        case CropKind.FreeCrop: ExecuteNativeCrop(hit); break;
                        case CropKind.Reset: ResetPictureCmd(hit); break;
                    }
                }
            }
            catch (Exception ex) { Log.Write("crop action 失败: " + choice.Label, ex); }

            // 形状 / 尺寸变了，下一个轮询周期重新给按钮定位
            _currentKey = "";
            _lastRect = Rectangle.Empty;
        }

        /// <summary>
        /// Excel 的裁剪动作分发。
        ///
        /// Excel 支持「裁剪为形状」的唯一办法是"自选图形 + 图片填充"（图片本身没有几何可改），
        /// 所以形状裁剪后选中的已经不是图片了。除"再裁一次形状"以外的动作，
        /// 都要先把图形还原成原图再执行 —— 和 WPS 里点完形状再选比例的表现一致。
        /// </summary>
        private void RunExcelCrop(PictureHit hit, CropChoice choice)
        {
            if (choice.Kind == CropKind.Shape)
            {
                ApplyExcelShapeCrop(hit, choice.Prst, choice.Label);
                return;
            }

            PictureHit pic = hit;
            if (hit.IsShapeImage)
            {
                pic = RestoreExcelPicture(hit);
                if (pic == null) return;
            }

            switch (choice.Kind)
            {
                case CropKind.Ratio: ApplyRatioCrop(pic, choice.Rw, choice.Rh, choice.Label); break;
                case CropKind.FreeCrop: ExecuteNativeCrop(pic); break;
                case CropKind.Reset: ResetPictureCmd(pic); break;
            }
        }

        private void SelectHit(PictureHit hit)
        {
            try
            {
                if (hit.IsInline && hit.Inline != null) hit.Inline.Select();
                else if (hit.Shape != null) hit.Shape.Select();
            }
            catch { }
        }

        /// <summary>
        /// 裁剪为形状：直接改 drawing 里的预设几何（prstGeom）再写回文档。
        ///
        /// 为什么不用 Shape.AutoShapeType：图片转出来的形状是 msoPicture，
        /// AutoShapeType 读回来永远是 -2（msoShapeMixed），赋值被静默丢弃（实测）。
        /// 改 XML 是唯一既能真正裁成形状、又保持"嵌入型"不变的办法 ——
        /// 实测 InsertXML 之后 InlineShapes 仍然是 1 个、Type 仍是 3（图片）。
        /// </summary>
        private void ApplyShapeCrop(PictureHit hit, string prst, string label)
        {
            bool floating = !hit.IsInline;
            int wrap = -1;
            dynamic shape = null;

            try
            {
                if (floating)
                {
                    shape = hit.Shape;
                    try { wrap = Convert.ToInt32(shape.WrapFormat.Type); } catch { }
                    // 浮动图取不到 WordOpenXML（实测返回空串），先转成嵌入型再改
                    shape = shape.ConvertToInlineShape();
                }
                else
                {
                    shape = hit.Inline;
                }

                dynamic range = shape.Range;
                int start = Convert.ToInt32(range.Start);
                string xml = Convert.ToString(range.WordOpenXML);

                if (string.IsNullOrEmpty(xml))
                {
                    Log.Write("crop shape: 取不到 XML, key=" + hit.Key);
                    MessageBox.Show(
                        "这张图片取不到可以修改的结构，没法裁剪成形状。\n\n" +
                        "可以改用 Word 自带的「图片格式 → 裁剪 → 裁剪为形状」。",
                        "裁剪为形状", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                string patched = SwapPrst(xml, prst);
                if (patched == null)
                {
                    Log.Write("crop shape: 没有 prstGeom, key=" + hit.Key);
                    MessageBox.Show(
                        "这张图片的形状信息不是标准格式，改不了。",
                        "裁剪为形状", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                range.InsertXML(patched);
                Log.Write("crop shape: " + label + " (" + prst + ") ok, key=" + hit.Key);

                // InsertXML 会把原来的 Range 内容换掉，选中状态随之丢失（实测 Selection.Type 变回 1）。
                // 必须把新图片重新选中，否则放大镜/裁剪按钮会一起消失，用户得再点一次图片。
                AfterInsert(start, floating, wrap);
            }
            catch (Exception ex)
            {
                Log.Write("crop shape: " + label + " 失败", ex);
                MessageBox.Show("裁剪成形状失败：" + ex.Message + "\n\n可以改用 Word 自带的裁剪。",
                                "裁剪为形状", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        /// <summary>
        /// InsertXML 之后收尾：找回被替换出来的图片。
        /// 嵌入型就重新选中它（保住按钮）；浮动图片则再变回浮动图片并还原环绕方式。
        /// </summary>
        private void AfterInsert(int start, bool floating, int wrap)
        {
            try
            {
                dynamic doc = _app.ActiveDocument;
                int n = Convert.ToInt32(doc.InlineShapes.Count);
                dynamic found = null;
                for (int i = 1; i <= n; i++)
                {
                    dynamic s = doc.InlineShapes.Item(i);
                    int st;
                    try { st = Convert.ToInt32(s.Range.Start); }
                    catch { continue; }
                    if (st == start) { found = s; break; }
                }

                if (found == null) { Log.Write("crop shape: 没找回替换后的图片（start=" + start + "）"); return; }

                if (!floating)
                {
                    found.Select();
                    return;
                }

                dynamic back = found.ConvertToShape();
                if (wrap >= 0)
                {
                    try { back.WrapFormat.Type = wrap; } catch { }
                }
                try { back.Select(); } catch { }
                Log.Write("crop shape: 已还原为浮动图片, wrap=" + wrap);
            }
            catch (Exception ex) { Log.Write("crop shape: 收尾失败", ex); }
        }

        /// <summary>把 XML 里第一个 prstGeom 的形状名换掉；没有就返回 null。</summary>
        private static string SwapPrst(string xml, string prst)
        {
            Match m = Regex.Match(xml, "<a:prstGeom\\s+prst=\"[^\"]*\"");
            if (!m.Success) return null;
            return xml.Substring(0, m.Index)
                 + "<a:prstGeom prst=\"" + prst + "\""
                 + xml.Substring(m.Index + m.Length);
        }

        // ------------------------------------------------------------ PowerPoint：裁剪为形状

        /// <summary>
        /// PowerPoint 的「裁剪为形状」：直接写 Shape.AutoShapeType。
        ///
        /// 三个宿主里最省事的一个 —— PowerPoint 的图片本身就是 <c>p:pic</c>，
        /// 里面的 <c>a:prstGeom</c> 可写，AutoShapeType 正好映射到它。
        /// 实测（tools/pptflow.cs）：
        ///   · 设成 9（椭圆）后 Shape.Type 仍然是 13（图片），没有变成形状对象；
        ///   · 导出的 PNG 四角 alpha=0、中心仍是原色 —— 是真裁掉了，不是假象；
        ///   · PictureFormat 仍然可用，所以「先裁形状再改比例」可以叠加，
        ///     不需要 Excel 那套「先把图形还原成原图」的绕行方案。
        /// </summary>
        private void ApplyPowerPointShapeCrop(PictureHit hit, string prst, string label)
        {
            try
            {
                dynamic shp = hit.Shape;
                if (shp == null) return;

                int ast = MsoShapeType(prst);
                shp.AutoShapeType = ast;
                Log.Write("ppt crop shape: " + label + " (" + prst + "/ast=" + ast + ") ok, id=" + hit.ShapeId);
            }
            catch (Exception ex)
            {
                Log.Write("ppt crop shape: " + label + " 失败", ex);
                MessageBox.Show(
                    "裁剪成形状失败：" + ex.Message + "\n\n" +
                    "可以改用 PowerPoint 自带的「图片格式 → 裁剪 → 裁剪为形状」。",
                    "裁剪为形状", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // ------------------------------------------------------------ Excel：裁剪为形状

        /// <summary>
        /// 「图片填充自选图形」的标记，写在 Shape.AlternativeText 里，格式
        /// <c>IZADDIN:SHAPE:&lt;prst&gt;|&lt;原图名字&gt;</c>。
        /// 有了原图名字，才能在「重设 / 换比例 / 换形状」时把藏在下面的原图找回来。
        /// </summary>
        private const string ShapeImageTag = "IZADDIN:SHAPE:";

        /// <summary>
        /// prstGeom 的预设几何名 → msoAutoShapeType 编号。Excel / PowerPoint 共用同一套枚举。
        ///
        /// 不是查文档抄的，是 tools/shapemap.cs 里对 1..250 逐个 AddShape 再解析
        /// drawing xml 的 prstGeom 反查出来的（见 tools/shapemap.txt）。
        /// PowerPoint 侧用 tools/pptflow.cs 抽验过 9（椭圆）和 92（五角星），与表一致。
        /// </summary>
        private static int MsoShapeType(string prst)
        {
            switch (prst)
            {
                case "roundRect": return 5;
                case "snipRoundRect": return 154;
                case "snip1Rect": return 155;
                case "snip2SameRect": return 156;
                case "snip2DiagRect": return 157;
                case "ellipse": return 9;
                case "triangle": return 7;
                case "rtTriangle": return 8;
                case "diamond": return 4;
                case "parallelogram": return 2;
                case "trapezoid": return 3;
                case "pentagon": return 12;
                case "hexagon": return 10;
                case "octagon": return 6;
                case "cross": return 11;          // Excel 里这个形状的 prst 叫 plus，外观一致
                case "plaque": return 28;
                case "chevron": return 52;
                case "homePlate": return 51;
                case "star4": return 91;
                case "star5": return 92;
                case "star6": return 147;
                case "star8": return 93;
                case "rightArrow": return 33;
                case "leftArrow": return 34;
                case "upArrow": return 35;
                case "downArrow": return 36;
                case "leftRightArrow": return 37;
                case "heart": return 21;
                case "cloud": return 179;
                case "moon": return 24;
                case "sun": return 23;
                case "teardrop": return 160;
                case "donut": return 18;
                default: return 1;                // rect
            }
        }

        private static bool IsShapeImage(dynamic shape)
        {
            try { return Convert.ToString(shape.AlternativeText).StartsWith(ShapeImageTag, StringComparison.Ordinal); }
            catch { return false; }
        }

        private static string MarkerOriginal(string alt)
        {
            if (string.IsNullOrEmpty(alt)) return "";
            int i = alt.LastIndexOf('|');
            return i >= 0 ? alt.Substring(i + 1) : "";
        }

        /// <summary>
        /// Excel 的「裁剪为形状」。
        ///
        /// Excel 的图片没有 Word 那种可以改的 prstGeom（x:pic 里也没有可写的几何接口），
        /// 所以唯一的真办法是：新建一个自选图形（形状对），把图片填进去（Fill.UserPicture），
        /// 尺寸和位置沿用原图，原图 Visible = msoFalse 藏起来 —— 留着给「重设形状和大小」还原。
        /// 实测（tools/xlprobe.cs）：Fill.Type 会变成 6（图片填充），且图片字节复用 xl/media 里
        /// 那一份，文件不会因此变大。
        /// </summary>
        private void ApplyExcelShapeCrop(PictureHit hit, string prst, string label)
        {
            string temp = null;
            try
            {
                dynamic src = hit.Shape;
                if (src == null) return;
                dynamic sheet = src.Parent;

                // 先把图片字节落到临时文件（Fill.UserPicture 只认文件路径），
                // 而且必须在动对象之前取 —— 后面会删掉旧形状。
                temp = WriteTempImageFile(hit);

                // 已经是形状填充图：先还原成原图，再按新形状重做，避免形状套形状
                if (hit.IsShapeImage)
                {
                    dynamic orig = FindOriginalPicture(src);
                    if (orig == null)
                    {
                        MessageBox.Show(
                            "这张图形已经找不到它对应的原图了（可能被删除或复制过）。\n\n" +
                            "请用 Ctrl+Z 撤销上一步，或重新插入图片。",
                            "裁剪为形状", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                    try { src.Delete(); } catch { }
                    try { orig.Visible = -1; } catch { }
                    try { orig.Select(); } catch { }
                    src = orig;
                }

                if (temp == null)
                {
                    MessageBox.Show(
                        "没能取到这张图片的原始数据，没法裁剪成形状。\n\n" +
                        "如果是链接到外部文件的图片，请先「断开链接」或重新插入一次。",
                        "裁剪为形状", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                double left = Convert.ToDouble(src.Left);
                double top = Convert.ToDouble(src.Top);
                double w = Convert.ToDouble(src.Width);
                double h = Convert.ToDouble(src.Height);
                string name = Convert.ToString(src.Name);

                dynamic sp = sheet.Shapes.AddShape(MsoShapeType(prst), left, top, w, h);
                sp.Fill.UserPicture(temp);
                try { sp.Line.Visible = 0; } catch { }
                try { sp.Placement = src.Placement; } catch { }
                try { sp.AlternativeText = ShapeImageTag + prst + "|" + name; } catch { }
                try { sp.Name = name + "#shape"; } catch { }
                try { src.Visible = 0; } catch { }          // 原图藏起来，供「重设」还原
                try { sp.Select(); } catch { }

                Log.Write("excel crop shape: " + label + " (" + prst + "/type=" + MsoShapeType(prst) +
                          ") ok, name=" + name + " key=" + hit.Key);
            }
            catch (Exception ex)
            {
                Log.Write("excel crop shape: " + label + " 失败", ex);
                MessageBox.Show("裁剪成形状失败：" + ex.Message, "裁剪为形状",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
                if (temp != null) { try { File.Delete(temp); } catch { } }
            }
        }

        /// <summary>从形状填充图上的标记找回被藏起来的原图（找不到就退化成"本表第一个隐藏图片"）。</summary>
        private dynamic FindOriginalPicture(dynamic shape)
        {
            try
            {
                dynamic sheet = shape.Parent;
                string orig = MarkerOriginal(Convert.ToString(shape.AlternativeText));
                if (!string.IsNullOrEmpty(orig))
                {
                    try
                    {
                        dynamic found = sheet.Shapes.Item(orig);
                        if (found != null) return found;
                    }
                    catch { }
                }

                int n = Convert.ToInt32(sheet.Shapes.Count);
                for (int i = 1; i <= n; i++)
                {
                    dynamic s = sheet.Shapes.Item(i);
                    if (Convert.ToInt32(s.Visible) != 0) continue;
                    int t = Convert.ToInt32(s.Type);
                    if (t == 11 || t == 13) return s;
                }
            }
            catch (Exception ex) { Log.Write("excel crop: 找回原图失败", ex); }
            return null;
        }

        /// <summary>把形状填充图还原成图片（删掉形状、原图恢复可见并选中），返回新的 PictureHit。</summary>
        private PictureHit RestoreExcelPicture(PictureHit hit)
        {
            try
            {
                dynamic shp = hit.Shape;
                dynamic orig = FindOriginalPicture(shp);
                if (orig == null)
                {
                    MessageBox.Show(
                        "这张图形已经找不到它对应的原图了（可能被删除或复制过）。\n\n" +
                        "请用 Ctrl+Z 撤销上一步，或重新插入图片。",
                        "裁剪图片", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    HideMagnifier();
                    return null;
                }

                try { shp.Delete(); } catch { }
                try { orig.Visible = -1; } catch { }
                try { orig.Select(); } catch { }

                PictureHit again = DetectExcelPicture();
                if (again != null)
                {
                    again.ScreenRect = GetExcelShapeRect(again.Shape);
                    return again;
                }

                // 兜底：自己拼一个（Selection 有时不会立刻刷新）
                PictureHit manual = new PictureHit();
                manual.Host = "Excel";
                manual.Shape = orig;
                manual.Name = Convert.ToString(orig.Name);
                manual.SheetName = hit.SheetName;
                manual.Key = "E:" + hit.SheetName + ":" + manual.Name;
                manual.ScreenRect = GetExcelShapeRect(orig);
                return manual;
            }
            catch (Exception ex)
            {
                Log.Write("excel crop: 还原成图片失败", ex);
                MessageBox.Show("还原成图片失败：" + ex.Message, "裁剪图片",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }
        }

        /// <summary>把图片原始字节写成临时文件（扩展名按真实格式给，避免重新编码掉画质）。</summary>
        private string WriteTempImageFile(PictureHit hit)
        {
            try
            {
                ImagePayload payload = Extract(hit);
                if (payload == null || payload.Bytes == null || payload.Bytes.Length == 0) return null;

                string path = Path.Combine(Path.GetTempPath(),
                                            "izshape_" + Guid.NewGuid().ToString("N") + ImageExt(payload.Bytes));
                File.WriteAllBytes(path, payload.Bytes);
                Log.Write("excel crop shape: 临时图 " + path + " (" + payload.Bytes.Length + " bytes, " + payload.Source + ")");
                return path;
            }
            catch (Exception ex) { Log.Write("excel crop shape: 写临时图失败", ex); return null; }
        }

        private static string ImageExt(byte[] b)
        {
            if (b == null || b.Length < 8) return ".png";
            if (b[0] == 0x89 && b[1] == 0x50) return ".png";
            if (b[0] == 0xFF && b[1] == 0xD8) return ".jpg";
            if (b[0] == 0x47 && b[1] == 0x49) return ".gif";
            if (b[0] == 0x42 && b[1] == 0x4D) return ".bmp";
            if (b[0] == 0x49 && b[1] == 0x49) return ".tif";
            if (b[0] == 0x4D && b[1] == 0x4D) return ".tif";
            return ".png";
        }

        /// <summary>
        /// 按比例裁剪：只动 PictureFormat 的四个裁剪值，不碰形状、不换对象，随时可以再改。
        ///
        /// 两个必须知道的事实（都实测确认过）：
        ///   1) CropLeft/Top 的单位不是当前尺寸，而是"原图 100% 尺寸的磅值"；
        ///   2) 裁掉 v 后可见宽度按比例缩小：W' = W * (1 - v / W0)。
        /// W0 不去猜 DPI，而是用一次探针（先裁 10% 看宽度掉多少）量出来，
        /// 再按 W0*(1 - 目标宽/当前宽) 计算，最后迭代两次把误差压到 0.2% 以内。
        /// </summary>
        private void ApplyRatioCrop(PictureHit hit, double rw, double rh, string label)
        {
            object raw = hit.IsInline ? (object)hit.Inline : (object)hit.Shape;
            dynamic shp = raw;
            dynamic pf = shp.PictureFormat;

            // 选比例 = 以原图为基准重新裁（和 WPS 一致，不会在旧裁剪上叠加）
            pf.CropLeft = 0.0;
            pf.CropRight = 0.0;
            pf.CropTop = 0.0;
            pf.CropBottom = 0.0;

            double Wf = Convert.ToDouble(shp.Width);
            double Hf = Convert.ToDouble(shp.Height);
            if (Wf < 1 || Hf < 1) return;

            if (rw <= 0 || rh <= 0)
            {
                Log.Write("crop ratio: 恢复原图比例 key=" + hit.Key +
                          " -> " + Wf.ToString("0.0") + "x" + Hf.ToString("0.0"));
                return;
            }

            double W0 = Wf;
            double Wp = Wf;
            double probe = Wf * 0.1;
            try
            {
                pf.CropLeft = probe;
                Wp = Convert.ToDouble(shp.Width);
                pf.CropLeft = 0.0;
                if (Wp > 0.01 && Wp < Wf * 0.999)
                    W0 = probe / (1.0 - Wp / Wf);
            }
            catch (Exception ex) { Log.Write("crop ratio: 探针失败，按当前尺寸估算", ex); }
            if (W0 < Wf * 0.5 || W0 > Wf * 20) W0 = Wf;
            double H0 = W0 * Hf / Wf;
            Log.Write("crop ratio: 探针 Wf=" + Wf.ToString("0.0") + " probe=" + probe.ToString("0.0") +
                      " Wp=" + Wp.ToString("0.0") + " -> W0=" + W0.ToString("0.0") + " key=" + hit.Key);

            double R = rw / rh;
            double cl = 0, cr = 0, ct = 0, cb = 0;

            for (int pass = 0; pass < 3; pass++)
            {
                double Wv = Convert.ToDouble(shp.Width);
                double Hv = Convert.ToDouble(shp.Height);
                if (Hv < 0.01) break;

                double cur = Wv / Hv;
                if (Math.Abs(cur - R) <= R * 0.0015) break;

                if (cur > R)
                {
                    // 太宽 → 左右各裁一半
                    double dc = (Wv - Hv * R) * W0 / Wf;
                    if (dc < 0.002) break;
                    cl += dc / 2.0;
                    cr += dc / 2.0;
                    pf.CropLeft = cl;
                    pf.CropRight = cr;
                }
                else
                {
                    // 太高 → 上下各裁一半
                    double dc = (Hv - Wv / R) * H0 / Hf;
                    if (dc < 0.002) break;
                    ct += dc / 2.0;
                    cb += dc / 2.0;
                    pf.CropTop = ct;
                    pf.CropBottom = cb;
                }
            }

            Log.Write("crop ratio: " + label + " -> " +
                      Convert.ToDouble(shp.Width).ToString("0.0") + "x" +
                      Convert.ToDouble(shp.Height).ToString("0.0") +
                      " (W0=" + W0.ToString("0.0") + ") key=" + hit.Key);
        }

        /// <summary>
        /// 重设形状和大小：用宿主自带的「重设图片和大小」（idMso = PictureResetAndSize）。
        /// Word 实测会把 prstGeom 还原成 rect、裁剪清零、尺寸回到默认（那张 4K 图回到 576x324）；
        /// Excel 实测裁剪清零、尺寸回到图片的原生大小（800x600px 的图 -> 600x450 磅）。
        /// </summary>
        private void ResetPictureCmd(PictureHit hit)
        {
            SelectHit(hit);
            try { _app.ActiveWindow.Activate(); } catch { }

            string before = "";
            try { before = Convert.ToDouble(hit.Shape != null ? hit.Shape.Width : hit.Inline.Width).ToString("0.0"); }
            catch { }

            try
            {
                _app.CommandBars.ExecuteMso("PictureResetAndSize");

                string after = "";
                try { after = Convert.ToDouble(hit.Shape != null ? hit.Shape.Width : hit.Inline.Width).ToString("0.0"); }
                catch { }
                Log.Write("crop reset: ok, key=" + hit.Key + " width " + before + " -> " + after);
            }
            catch (Exception ex)
            {
                Log.Write("crop reset: 失败 :: " + ex.Message + CropState());
                MessageBox.Show(
                    "没能重设图片。\n\n" + CropHint() + "\n\n诊断信息：\n" + CropState().Trim(),
                    "重设形状和大小", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        /// <summary>
        /// 自由裁剪：触发 Word 自带的「图片格式 → 裁剪」Ribbon 命令（idMso = PictureCrop），
        /// 进入 Word 原生裁剪状态（可拖四边/四角、可指定比例）。
        ///
        /// 这个命令能否执行取决于 Word 当时的上下文，单看返回值看不出来，所以：
        ///   1) 失败先自动重试一次（先把 Word 窗口/文档激活、重新选中图片再试）；
        ///   2) 仍然失败时，把 Word 的真实状态（enabled/readOnly/final/prot/视图）记进日志，
        ///      并在弹窗里显示出来 —— 远程定位全靠这一行。
        /// </summary>
        private void ExecuteNativeCrop(PictureHit hit)
        {
            // Ribbon 命令作用在「当前选中对象」上；按钮窗口是 NOACTIVATE 的，
            // 选中状态本来就不会丢，这里再选一次纯属保险。
            SelectHit(hit);

            // 第一次就先激活一次：Ribbon 命令是按「活动窗口」的上下文决定能不能执行的，
            // 窗口处于非活动状态时 ExecuteMso 会直接抛 E_FAIL。
            try { _app.ActiveWindow.Activate(); } catch (Exception ex) { Log.Write("crop: 激活窗口失败", ex); }

            for (int attempt = 1; attempt <= 2; attempt++)
            {
                try
                {
                    if (attempt == 2)
                    {
                        // 第一次失败多是「Word 窗口不活动 / 上下文没刷新」，激活后再来一遍。
                        try { _app.Activate(); } catch { }
                        try { _app.ActiveWindow.Activate(); } catch { }
                        SelectHit(hit);
                        System.Threading.Thread.Sleep(80);
                    }

                    _app.CommandBars.ExecuteMso("PictureCrop");
                    Log.Write("crop: ok (attempt=" + attempt + ") key=" + hit.Key + CropState());
                    return;
                }
                catch (Exception ex)
                {
                    Log.Write("crop: 第 " + attempt + " 次失败 :: " + ex.Message + CropState());
                }
            }

            string state = CropState().Trim();
            Log.Write("crop: 放弃 -> " + state);
            MessageBox.Show(
                _host + " 拒绝了裁剪命令。\n\n" + CropHint() + "\n\n诊断信息：\n" + state,
                "裁剪图片", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        /// <summary>裁剪失败时按宿主的实际情况说人话，而不是笼统地"请再试一次"。</summary>
        private string CropHint()
        {
            try
            {
                if (_host == "Excel")
                {
                    dynamic wb = _app.ActiveWorkbook;
                    if (Convert.ToBoolean(wb.ReadOnly))
                        return "这个工作簿是「只读」状态，Excel 会禁用裁剪。\n" +
                               "用「文件 → 另存为」存一个副本，在副本里裁剪。";
                    dynamic sheet = _app.ActiveSheet;
                    if (Convert.ToBoolean(sheet.ProtectContents))
                        return "当前工作表处于保护状态，裁剪被禁用。\n请先到「审阅 → 撤消工作表保护」。";
                    return "请确认图片处于选中状态，或直接用「图片格式 → 裁剪」。";
                }

                if (_host == "PowerPoint")
                {
                    dynamic prs = _app.ActivePresentation;
                    if (Convert.ToBoolean(prs.ReadOnly))
                        return "这个演示文稿是「只读」状态，PowerPoint 会禁用裁剪。\n" +
                               "用「文件 → 另存为」存一个副本，在副本里裁剪。";
                    if (Convert.ToBoolean(prs.Final))
                        return "演示文稿被标记为「最终版本」，PowerPoint 按只读处理。\n请到「文件 → 信息 → 仍然编辑」。";
                    return "请确认图片处于选中状态（不要是组合里的），或直接用「图片格式 → 裁剪」。";
                }

                dynamic doc = _app.ActiveDocument;
                if (Convert.ToBoolean(doc.ReadOnly))
                    return "这个文档是「只读」状态（标题栏带“只读”），Word 会禁用裁剪。\n" +
                           "用「文件 → 另存为」存一个副本，在副本里裁剪；或直接在可编辑的文档里试。";
                if (Convert.ToInt32(doc.ProtectionType) != -1)
                    return "文档处于保护状态（限制编辑），裁剪被禁用。\n请先到「审阅 → 限制编辑 → 停止保护」。";
                if (Convert.ToBoolean(doc.Final))
                    return "文档被标记为「最终版本」，Word 按只读处理。\n请到「文件 → 信息 → 仍然编辑」。";
            }
            catch { }
            return "请确认图片处于选中状态，或直接用「图片格式 → 裁剪」。";
        }

        /// <summary>把宿主的关键状态一次记全（日志里跟在那行错误后面）。</summary>
        private string CropState()
        {
            StringBuilder sb = new StringBuilder(" | enabled=");
            try { sb.Append(Convert.ToString(_app.CommandBars.GetEnabledMso("PictureCrop"))); }
            catch (Exception ex) { sb.Append("?(").Append(ex.Message).Append(')'); }

            try
            {
                if (_host == "Excel")
                {
                    dynamic wb = _app.ActiveWorkbook;
                    dynamic sheet = _app.ActiveSheet;
                    sb.Append(" wb=").Append(Convert.ToString(wb.Name));
                    sb.Append(" readOnly=").Append(Convert.ToString(wb.ReadOnly));
                    sb.Append(" prot=").Append(Convert.ToString(wb.ProtectStructure));
                    sb.Append(" sheetProt=").Append(Convert.ToString(sheet.ProtectContents));
                    sb.Append(" windows=").Append(Convert.ToString(_app.Windows.Count));
                    return sb.ToString();
                }

                if (_host == "PowerPoint")
                {
                    dynamic prs = _app.ActivePresentation;
                    sb.Append(" ppt=").Append(Convert.ToString(prs.Name));
                    sb.Append(" readOnly=").Append(Convert.ToString(prs.ReadOnly));
                    sb.Append(" final=").Append(Convert.ToString(prs.Final));
                    sb.Append(" slides=").Append(Convert.ToString(prs.Slides.Count));
                    sb.Append(" viewType=").Append(Convert.ToString(_app.ActiveWindow.ViewType));
                    sb.Append(" windows=").Append(Convert.ToString(_app.Windows.Count));
                    return sb.ToString();
                }

                dynamic doc = _app.ActiveDocument;
                sb.Append(" doc=").Append(Convert.ToString(doc.Name));
                sb.Append(" readOnly=").Append(Convert.ToString(doc.ReadOnly));
                sb.Append(" final=").Append(Convert.ToString(doc.Final));
                sb.Append(" prot=").Append(Convert.ToString(doc.ProtectionType));
                sb.Append(" view=").Append(Convert.ToString(_app.ActiveWindow.View.Type));
                sb.Append(" windows=").Append(Convert.ToString(_app.Windows.Count));
            }
            catch (Exception ex) { sb.Append(" state=?(").Append(ex.Message).Append(')'); }
            return sb.ToString();
        }

        private void ClosePreview()
        {
            try
            {
                if (_preview != null && !_preview.IsDisposed)
                {
                    _preview.Close();
                    _preview.Dispose();
                }
            }
            catch { }
            _preview = null;
        }

        // ------------------------------------------------------------ 取图

        private ImagePayload Extract(PictureHit hit)
        {
            // 1) Word 嵌入式图片：读 Flat OPC 里的原始字节（最准确）
            if (hit.IsInline && hit.Inline != null)
            {
                try
                {
                    string flat = Convert.ToString(hit.Inline.Range.WordOpenXML);
                    byte[] bytes = Ooxml.FromFlatOpc(flat);
                    if (bytes != null && bytes.Length > 128)
                        return new ImagePayload(bytes, "OOXML 原始字节");
                }
                catch (Exception ex) { Log.Write("Flat OPC 取图失败，走后备方案", ex); }
            }

            // 2) Excel：解包 xlsx，从 xl/media 取原图（最准确）
            if (hit.Host == "Excel")
            {
                try
                {
                    byte[] bytes = ExtractFromExcelPackage(hit);
                    if (bytes != null && bytes.Length > 128)
                        return new ImagePayload(bytes, "xlsx 包内原图");
                }
                catch (Exception ex) { Log.Write("xlsx 解包取图失败", ex); }
            }

            // 2.5) PowerPoint：解包 pptx，从 ppt/media 取原图
            if (hit.Host == "PowerPoint")
            {
                try
                {
                    byte[] bytes = ExtractFromPowerPointPackage(hit);
                    if (bytes != null && bytes.Length > 128)
                        return new ImagePayload(bytes, "pptx 包内原图");
                }
                catch (Exception ex) { Log.Write("pptx 解包取图失败", ex); }
            }

            // 3) 剪贴板：优先原始格式流（Word/Excel 复制图片时通常带 PNG/JFIF 原始字节）
            try
            {
                if (hit.Host == "Excel" && hit.Shape != null)
                {
                    hit.Shape.Copy();
                }
                else
                {
                    SelectTarget(hit);
                    _app.Selection.Copy();
                }
                System.Threading.Thread.Sleep(80);

                byte[] raw = ClipboardHelper.TryRawImageBytes();
                if (raw != null) return new ImagePayload(raw, "剪贴板原图字节");

                Bitmap rendered = ClipboardHelper.TryRenderedImage();
                if (rendered != null)
                {
                    byte[] png = ClipboardHelper.ImageToPngBytes(rendered);
                    rendered.Dispose();
                    if (png != null) return new ImagePayload(png, "剪贴板渲染位图（可能非原图）");
                }
            }
            catch (Exception ex) { Log.Write("剪贴板取图失败", ex); }
            finally
            {
                try { if (hit.Host == "Excel") _app.CutCopyMode = false; } catch { }
            }

            return null;
        }

        private void SelectTarget(PictureHit hit)
        {
            try
            {
                if (hit.IsInline && hit.Inline != null) hit.Inline.Select();
                else if (hit.Shape != null) hit.Shape.Select();
            }
            catch (Exception ex) { Log.Write("重新选中图片失败", ex); }
        }

        private byte[] ExtractFromExcelPackage(PictureHit hit)
        {
            dynamic wb = _app.ActiveWorkbook;
            if (wb == null) return null;

            string temp = Path.Combine(Path.GetTempPath(), "izp_" + Guid.NewGuid().ToString("N") + ".xlsx");
            try
            {
                wb.SaveCopyAs(temp);
                if (!File.Exists(temp)) return null;
                return Ooxml.FromExcelPackage(temp, hit.Name, hit.SheetName);
            }
            finally
            {
                try { if (File.Exists(temp)) File.Delete(temp); } catch { }
            }
        }

        /// <summary>
        /// PowerPoint 取原图：把当前演示文稿 SaveCopyAs 成临时 .pptx，解包后按 Shape.Id 找 media。
        ///
        /// SaveCopyAs 对"从未保存过的新演示文稿"同样可用（实测 Path 为空也能生成完整包），
        /// 不会动用户的文件、也不会留下未保存标记。
        /// </summary>
        private byte[] ExtractFromPowerPointPackage(PictureHit hit)
        {
            dynamic prs = null;
            try { prs = _app.ActivePresentation; } catch { }
            if (prs == null) return null;

            string temp = Path.Combine(Path.GetTempPath(), "izp_" + Guid.NewGuid().ToString("N") + ".pptx");
            try
            {
                prs.SaveCopyAs(temp);
                if (!File.Exists(temp)) return null;

                byte[] bytes = Ooxml.FromPowerPointPackage(temp, hit.ShapeId, hit.SlideIndex);
                Log.Write("pptx: 包内取图 id=" + hit.ShapeId + " slide=" + hit.SlideIndex + " -> " +
                          (bytes == null ? "未命中" : bytes.Length + " bytes"));
                return bytes;
            }
            finally
            {
                try { if (File.Exists(temp)) File.Delete(temp); } catch { }
            }
        }

        private static Bitmap Decode(byte[] bytes)
        {
            if (bytes == null || bytes.Length == 0) return null;
            try
            {
                using (MemoryStream ms = new MemoryStream(bytes))
                using (Image img = Image.FromStream(ms, true, true))
                {
                    return new Bitmap(img);
                }
            }
            catch (Exception ex)
            {
                Log.Write("图片解码失败", ex);
                return null;
            }
        }
    }

    #endregion
}
