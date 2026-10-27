// 端到端：在真实 PowerPoint 里跑加载项，验证放大镜/裁剪/旋转三个按钮
//   1) 加载项到底有没有被 PowerPoint 拉起来
//   2) 按钮算出来的 rect 跟真实截图位置对不对得上（用 RangeFromPoint 反量）
//   3) 旋转 90° 之后按钮有没有跟着挪
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

internal static class PptLive
{
    [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr h, int cmd);
    [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
    [DllImport("user32.dll")] private static extern bool AttachThreadInput(uint a, uint b, bool f);
    [DllImport("user32.dll")] private static extern bool BringWindowToTop(IntPtr h);
    [DllImport("user32.dll")] private static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint flags);
    [DllImport("user32.dll")] private static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] private static extern void mouse_event(uint flags, int dx, int dy, uint data, IntPtr extra);
    [DllImport("user32.dll")] private static extern bool EnumWindows(EnumProc cb, IntPtr p);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] private static extern int GetWindowText(IntPtr h, StringBuilder s, int max);
    [DllImport("user32.dll")] private static extern int GetClassName(IntPtr h, StringBuilder s, int max);
    [DllImport("user32.dll")] private static extern bool PostMessage(IntPtr h, uint msg, IntPtr w, IntPtr l);

    private delegate bool EnumProc(IntPtr h, IntPtr p);

    private static Point ParsePoint(string s, string key)
    {
        int i = s.IndexOf(key);
        string seg = s.Substring(i + key.Length + 1);      // 跳过 '{'
        int j = seg.IndexOf('}');
        string body = seg.Substring(0, j);
        string[] kv = body.Split(',');
        return new Point(int.Parse(kv[0].Split('=')[1]), int.Parse(kv[1].Split('=')[1]));
    }

    private static string WindowText(IntPtr h)
    {
        StringBuilder sb = new StringBuilder(512);
        GetWindowText(h, sb, sb.Capacity);
        return sb.ToString();
    }

    /// <summary>点一下 (x,y)，然后在 timeout 内找标题里带「原图」的可见窗口。</summary>
    private static IntPtr ClickAndFindPreview(int x, int y, int timeoutMs)
    {
        SetCursorPos(x, y);
        Thread.Sleep(120);
        mouse_event(0x0002, 0, 0, 0, IntPtr.Zero);   // LEFTDOWN
        Thread.Sleep(60);
        mouse_event(0x0004, 0, 0, 0, IntPtr.Zero);   // LEFTUP

        IntPtr found = IntPtr.Zero;
        for (int t = 0; t < timeoutMs; t += 250)
        {
            Thread.Sleep(250);
            EnumWindows(delegate(IntPtr h, IntPtr p)
            {
                if (!IsWindowVisible(h)) return true;
                string txt = WindowText(h);
                if (txt.IndexOf("原图", StringComparison.Ordinal) < 0) return true;
                StringBuilder cls = new StringBuilder(256);
                GetClassName(h, cls, cls.Capacity);
                if (cls.ToString().IndexOf("WindowsForms", StringComparison.Ordinal) < 0) return true;
                found = h;
                return false;
            }, IntPtr.Zero);
            if (found != IntPtr.Zero) return found;
        }
        return IntPtr.Zero;
    }


    private static void P(string k, object v) { Console.WriteLine("  " + k + " = " + v); }

    private static IntPtr _hwnd = IntPtr.Zero;
    private static dynamic _win;
    private static string _name = "";

    private static void Front()
    {
        if (_hwnd == IntPtr.Zero) return;
        IntPtr fg = GetForegroundWindow();
        uint pid;
        uint fgThread = fg == IntPtr.Zero ? 0 : GetWindowThreadProcessId(fg, out pid);
        uint myThread = GetCurrentThreadId();
        bool attached = false;
        if (fgThread != 0 && fgThread != myThread) attached = AttachThreadInput(fgThread, myThread, true);
        try
        {
            // 注意：这里**不能**把 PowerPoint 设成 HWND_TOPMOST ——
            // 按钮窗口自己就是 TopMost，PowerPoint 一旦也 topmost 且后设置，
            // 就会把按钮盖住，后面的真实点击全落到 PowerPoint 上。
            // 只需要让它成为前台窗口（加载项的轮询要求前台是本进程）。
            SetWindowPos(_hwnd, new IntPtr(-2), 0, 0, 0, 0, 0x0043);   // HWND_NOTOPMOST
            BringWindowToTop(_hwnd);
            ShowWindow(_hwnd, 3);
            SetForegroundWindow(_hwnd);
        }
        finally
        {
            if (attached) AttachThreadInput(fgThread, myThread, false);
        }
        Thread.Sleep(200);
    }

    private static bool Hit(int x, int y)
    {
        try
        {
            dynamic shp = _win.RangeFromPoint((float)x, (float)y);
            if (shp == null) return false;
            return string.Equals(Convert.ToString(shp.Name), _name, StringComparison.Ordinal);
        }
        catch { return false; }
    }

    private static Rectangle Measure(int cx, int cy)
    {
        if (!Hit(cx, cy)) return Rectangle.Empty;
        int l = cx; while (l > 1 && Hit(l - 1, cy)) l--;
        int r = cx; while (r < 3999 && Hit(r + 1, cy)) r++;
        int t = cy; while (t > 1 && Hit(cx, t - 1)) t--;
        int b = cy; while (b < 2999 && Hit(cx, b + 1)) b++;
        return Rectangle.FromLTRB(l, t, r + 1, b + 1);
    }

    private static Rectangle ParseRect(string s)
    {
        // 日志格式: rect={X=731,Y=481,Width=800,Height=450}
        try
        {
            int x = s.IndexOf("{X=") + 3, y = s.IndexOf(",Y=") + 3, w = s.IndexOf(",Width=") + 7, h = s.IndexOf(",Height=") + 8;
            int xe = s.IndexOf(',', x), ye = s.IndexOf(',', y), we = s.IndexOf(',', w), he = s.IndexOf('}', h);
            return new Rectangle(int.Parse(s.Substring(x, xe - x)), int.Parse(s.Substring(y, ye - y)),
                                 int.Parse(s.Substring(w, we - w)), int.Parse(s.Substring(h, he - h)));
        }
        catch { return Rectangle.Empty; }
    }

    private static string LogPath()
    {
        return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                            "ImageZoomAddin", "addin.log");
    }

    private static long _logPos;
    private static List<string> NewLines()
    {
        List<string> res = new List<string>();
        try
        {
            if (!File.Exists(LogPath())) return res;
            using (FileStream fs = new FileStream(LogPath(), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                if (fs.Length <= _logPos) return res;
                fs.Seek(_logPos, SeekOrigin.Begin);
                using (StreamReader sr = new StreamReader(fs, Encoding.UTF8))
                {
                    string line;
                    while ((line = sr.ReadLine()) != null) res.Add(line);
                }
                _logPos = fs.Length;
            }
        }
        catch { }
        return res;
    }

    /// <summary>等加载项打印出一行包含 needle 的日志（最多 timeoutMs 毫秒）。</summary>
    private static string WaitLog(string needle, int timeoutMs)
    {
        int step = 200;
        for (int t = 0; t < timeoutMs; t += step)
        {
            Front();
            foreach (string l in NewLines())
                if (l.Contains(needle)) return l;
            Thread.Sleep(step);
        }
        return null;
    }

    private static string LastMagnifierLine()
    {
        string found = null;
        foreach (string l in NewLines())
            if (l.Contains("magnifier:")) found = l;
        return found;
    }

    [STAThread]
    private static int Main()
    {
        string pngRed = Path.Combine(Path.GetTempPath(), "izlive_red.png");
        string pngBlue = Path.Combine(Path.GetTempPath(), "izlive_blue.png");
        using (Bitmap b = new Bitmap(800, 450)) using (Graphics g = Graphics.FromImage(b)) { g.Clear(Color.Magenta); b.Save(pngRed, ImageFormat.Png); }
        using (Bitmap b = new Bitmap(640, 360)) using (Graphics g = Graphics.FromImage(b)) { g.Clear(Color.Cyan); b.Save(pngBlue, ImageFormat.Png); }

        Console.WriteLine("log = " + LogPath());
        try { if (File.Exists(LogPath())) _logPos = new FileInfo(LogPath()).Length; } catch { }
        _logPos = Math.Max(0, _logPos);

        Type t = Type.GetTypeFromProgID("PowerPoint.Application");
        dynamic app = Activator.CreateInstance(t);
        dynamic prs = null;
        try
        {
            app.Visible = true;
            prs = app.Presentations.Add();
            dynamic slide = prs.Slides.Add(1, 12);
            _win = app.ActiveWindow;

            dynamic pic = slide.Shapes.AddPicture(pngRed, 0, -1, 200, 200, 600, 337.5);
            pic.Name = "livepic";
            _name = "livepic";

            for (int i = 0; i < 20; i++)
            {
                foreach (System.Diagnostics.Process pr in System.Diagnostics.Process.GetProcessesByName("POWERPNT"))
                    if (pr.MainWindowHandle != IntPtr.Zero) { _hwnd = pr.MainWindowHandle; break; }
                if (_hwnd != IntPtr.Zero) break;
                Thread.Sleep(300);
            }
            try { app.WindowState = 3; } catch { }
            ShowWindow(_hwnd, 3);
            Front();

            pic.Select();
            Thread.Sleep(400);

            Console.WriteLine("==== 1. 加载项是否被拉起 ====");
            string line = WaitLog("magnifier:", 12000);
            if (line == null)
            {
                Console.WriteLine("  !! 12 秒内没有 magnifier 日志，后面全部跳过");
                foreach (string l in NewLines()) Console.WriteLine("    | " + l);
                return 2;
            }
            foreach (string l in NewLines())
                if (l.Contains("OnConnection") || l.Contains("StartPolling") || l.Contains("diag#"))
                    Console.WriteLine("    | " + l);
            Console.WriteLine("  magnifier 行: " + line);

            Console.WriteLine("==== 2. 按钮位置 vs 真实位置 ====");
            string mag = line;
            int idx = mag.IndexOf("rect=");
            Rectangle calc = ParseRect(mag.Substring(idx + 5));
            P("加载项算的 rect", calc.X + "," + calc.Y + " " + calc.Width + "x" + calc.Height);

            Rectangle real = Measure(calc.X + calc.Width / 2, calc.Y + calc.Height / 2);
            P("RangeFromPoint 量的", real.IsEmpty ? "(miss)" : real.X + "," + real.Y + " " + real.Width + "x" + real.Height);
            if (!real.IsEmpty)
            {
                // RangeFromPoint 在选中状态下每边外扩 8px（已知偏差）
                int dx = calc.X - (real.X + 8), dy = calc.Y - (real.Y + 8);
                int dw = calc.Width - (real.Width - 16), dh = calc.Height - (real.Height - 16);
                P("校准后误差", "dx=" + dx + " dy=" + dy + " dw=" + dw + " dh=" + dh +
                    (Math.Abs(dx) <= 3 && Math.Abs(dy) <= 3 && Math.Abs(dw) <= 4 && Math.Abs(dh) <= 4 ? "  OK" : "  *** 偏了 ***"));
            }

            // 按钮到底在不在那个位置：截一小块下来看看
            Console.WriteLine("==== 3. 按钮窗口处截图像素 ====");
            DumpButtonArea(mag);

            Console.WriteLine("==== 4. 旋转 90° 后按钮跟随 ====");
            pic.Rotation = 90;
            Thread.Sleep(300);
            pic.Select();
            string mag2 = WaitLog("magnifier:", 6000);
            // 旋转后 key 不变、rect 变了，一定会有新的 magnifier 行
            string lastLine = null;
            for (int i = 0; i < 25; i++)
            {
                Front();
                string l = LastMagnifierLine();
                if (l != null) lastLine = l;
                Thread.Sleep(150);
            }
            if (lastLine != null && lastLine != mag)
            {
                Console.WriteLine("  旋转后: " + lastLine);
                Rectangle c2 = ParseRect(lastLine.Substring(lastLine.IndexOf("rect=") + 5));
                Rectangle r2 = Measure(c2.X + c2.Width / 2, c2.Y + c2.Height / 2);
                P("calc", c2.X + "," + c2.Y + " " + c2.Width + "x" + c2.Height);
                P("real", r2.IsEmpty ? "(miss)" : r2.X + "," + r2.Y + " " + r2.Width + "x" + r2.Height);
                if (!r2.IsEmpty)
                {
                    P("外接矩形", "应约 " + (337.5 / 600.0 * c2.Width).ToString("0") + "x" + c2.Height +
                                  "，实际 " + (r2.Width - 16) + "x" + (r2.Height - 16));
                }
            }
            else
            {
                Console.WriteLine("  !! 旋转后没有出现新的 magnifier 行（按钮可能没跟随）");
            }
            pic.Rotation = 0;

            // ---------- 4.5 真点一下放大镜按钮 ----------
            Console.WriteLine("==== 4.5 点击放大镜 -> 预览窗口 ====");
            Thread.Sleep(600);
            Front();
            string cur = null;
            for (int i = 0; i < 20; i++) { cur = LastMagnifierLine(); if (cur != null) break; Thread.Sleep(150); }
            if (cur == null) { Console.WriteLine("  没有按钮日志，跳过"); }
            else
            {
                Point bp = ParsePoint(cur, "mag=");
                P("按钮左上角", bp.X + "," + bp.Y);
                IntPtr dlg = ClickAndFindPreview(bp.X + 15, bp.Y + 15, 4000);
                if (dlg == IntPtr.Zero)
                {
                    Console.WriteLine("  !! 点击后没找到预览窗口，列出当前可见顶层窗口：");
                    EnumWindows(delegate(IntPtr h, IntPtr p)
                    {
                        if (!IsWindowVisible(h)) return true;
                        string txt = WindowText(h);
                        if (txt.Length == 0) return true;
                        StringBuilder cls = new StringBuilder(256);
                        GetClassName(h, cls, cls.Capacity);
                        Console.WriteLine("     [" + cls + "] " + txt);
                        return true;
                    }, IntPtr.Zero);
                }
                else
                {
                    Console.WriteLine("  找到预览窗口: " + WindowText(dlg));
                    PostMessage(dlg, 0x0010, IntPtr.Zero, IntPtr.Zero);   // WM_CLOSE
                    Thread.Sleep(400);
                    Console.WriteLine("  已关闭预览窗口");
                }
            }

            Console.WriteLine("==== 5. 取消选中 ====");
            Thread.Sleep(300);
            try { slide.Shapes.Item(1).Select(); } catch { }
            // 点空白处：用 RangeFromPoint 拿不到"无选中"，直接切到第 2 页再回来
            try
            {
                prs.Slides.Add(2, 12);
                NewLines();                       // 先清空，只看切页之后新增的
                _win.View.GotoSlide(2);
                Thread.Sleep(1200);
                NewLines();
                Thread.Sleep(600);
                bool anyNew = false;
                foreach (string l in NewLines()) if (l.Contains("magnifier:")) { anyNew = true; Console.WriteLine("    | " + l); }
                P("切到第 2 页后还有新 magnifier 行", anyNew ? "有（不该有）" : "无（正确）");
            }
            catch (Exception ex) { P("slide switch", ex.Message); }
        }
        catch (Exception ex) { Console.WriteLine("ERR " + ex); }
        finally
        {
            try { if (prs != null) prs.Close(); } catch { }
            try { app.Quit(); } catch { }
            try { Marshal.ReleaseComObject(app); } catch { }
        }
        return 0;
    }

    private static void DumpButtonArea(string magLine)
    {
        // 日志里 mag= 是按钮左上角
        try
        {
            int i = magLine.IndexOf("mag=");
            string seg = magLine.Substring(i + 4);
            int j = seg.IndexOf('}');
            // 格式 {X=1071,Y=781}
            string body = seg.Substring(1, j - 1);
            string[] kv = body.Split(',');
            int bx = int.Parse(kv[0].Split('=')[1]);
            int by = int.Parse(kv[1].Split('=')[1]);
            P("mag 按钮左上角", bx + "," + by);

            int w = 30, h = 30;
            using (Bitmap shot = new Bitmap(w + 8, h + 8))
            using (Graphics g = Graphics.FromImage(shot))
            {
                g.CopyFromScreen(bx - 4, by - 4, 0, 0, new Size(w + 8, h + 8));
                // 统计非背景像素：按钮底色接近白 / 边框灰
                int light = 0, total = 0;
                for (int y = 0; y < shot.Height; y++)
                    for (int x = 0; x < shot.Width; x++)
                    {
                        Color c = shot.GetPixel(x, y);
                        total++;
                        if (c.R > 235 && c.G > 235 && c.B > 235) light++;
                    }
                P("浅色像素占比", (light * 100 / total) + "%  (" + light + "/" + total + ")");
                string save = Path.Combine(Path.GetTempPath(), "izlive_btn.png");
                shot.Save(save, ImageFormat.Png);
                P("截图", save);
            }
        }
        catch (Exception ex) { P("DumpButtonArea", ex.Message); }
    }
}
