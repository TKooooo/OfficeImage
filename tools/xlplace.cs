// Excel 按钮定位"地面对照"：让加载项自己算，截屏找纯品红图片的真实边界，两边对账。
//   - 刚插入 / 放大 / 平移 / 滚动工作表 / 改工作表缩放 五种情况各验一次
// 通过条件：放大镜左边 = 图片右边 + 6px，按钮组底边 = 图片底边 - 8px（容差 4px）
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

internal static class XlPlace
{
    [DllImport("user32.dll")] private static extern int GetSystemMetrics(int i);
    [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr h, int cmd);

    private static Type _connect;
    private static object _conn;
    private static dynamic _xl;
    private static dynamic _pic;
    private static int _fail;

    [STAThread]
    private static int Main(string[] args)
    {
        Assembly asm = Assembly.LoadFrom(args[0]);
        string img = Path.Combine(Path.GetTempPath(), "izplace.png");
        using (Bitmap bmp = new Bitmap(800, 450))
        using (Graphics g = Graphics.FromImage(bmp)) { g.Clear(Color.Magenta); bmp.Save(img, ImageFormat.Png); }

        _connect = asm.GetType("ImageZoomAddin.Connect");
        _conn = Activator.CreateInstance(_connect);

        _xl = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
        SetField("_app", (object)_xl);
        SetField("_host", "Excel");
        _xl.DisplayAlerts = false;
        _xl.Visible = true;
        _xl.WindowState = -4137;                  // xlMaximized
        dynamic wb = _xl.Workbooks.Add();
        dynamic ws = wb.Worksheets.Item(1);
        dynamic win = _xl.ActiveWindow;

        _pic = ws.Shapes.AddPicture(img, 0, -1, 120, 120, 420, 236);
        _pic.Name = "place_img";
        _pic.Select();

        // 等 Excel 真的画出来（PointsToScreenPixels 在窗口就绪前返回 0）
        for (int i = 0; i < 40; i++)
        {
            try
            {
                if ((IntPtr)_xl.Hwnd != IntPtr.Zero) ShowWindow((IntPtr)_xl.Hwnd, 3);
                SetForegroundWindow((IntPtr)_xl.Hwnd);
                if (Convert.ToInt32(win.PointsToScreenPixelsX(0)) != 0) break;
            }
            catch { }
            Thread.Sleep(250);
        }
        Thread.Sleep(1500);

        Check("刚插入 420x236pt");
        _pic.Width = 520; _pic.Height = 292; _pic.Select(); Thread.Sleep(500);
        Check("放大到 520x292");
        _pic.Left = 200; _pic.Top = 200; _pic.Select(); Thread.Sleep(500);
        Check("平移到 200,200");
        win.ScrollRow = 5; win.ScrollColumn = 2; Thread.Sleep(600);
        Check("滚动到 5行/2列");
        win.Zoom = 60; Thread.Sleep(1000);
        Check("工作表缩放到 60%");
        win.Zoom = 100; win.ScrollRow = 1; win.ScrollColumn = 1; Thread.Sleep(800);
        Check("回到 100% 且滚回原点");

        try { wb.Close(false); } catch { }
        try { _xl.Quit(); } catch { }
        Console.WriteLine(_fail == 0 ? "ALL PASS" : ("FAILED " + _fail));
        return _fail == 0 ? 0 : 1;
    }

    private static void Check(string tag)
    {
        try { _pic.Select(); } catch { }
        object hit = Invoke("DetectPicture");
        if (hit == null) { Console.WriteLine("[" + tag + "] hit=null  <<< FAIL"); _fail++; return; }
        Rectangle calc = (Rectangle)FieldObj(hit, "ScreenRect");
        Invoke("ShowMagnifier", hit);
        Pump();

        // 再读一次：如果和上面不一样，说明是"Excel 还在动"的时序问题，不是换算问题
        object hit2 = Invoke("DetectPicture");
        Rectangle calc2 = Rectangle.Empty;
        if (hit2 != null) { object o = FieldObj(hit2, "ScreenRect"); if (o is Rectangle) calc2 = (Rectangle)o; }
        Invoke("ShowMagnifier", hit2);
        Pump();

        Rectangle real = Rectangle.Empty;
        // Excel 重绘中间态会让截屏抓到"整块"颜色，用宽高比过滤掉，最多重试 4 次
        for (int attempt = 0; attempt < 4; attempt++)
        {
            Rectangle r = FindMagenta();
            double want = (double)calc.Width / Math.Max(1, calc.Height);
            double got = r.IsEmpty ? 0 : (double)r.Width / r.Height;
            if (!r.IsEmpty && Math.Abs(want - got) < 0.06 && Math.Abs(r.Width - calc.Width) <= 6) { real = r; break; }
            Thread.Sleep(350);
        }
        Rectangle mag = Bounds("_magnifier");
        Rectangle crop = Bounds("_crop");

        // 布局约定：放大镜左边 = 图片右边 + 6；放大镜底边 = 图片底边 - 8；裁剪在放大镜下方 +34
        int dLeft = mag.X - (real.Right + 6);
        int dBottom = mag.Bottom - (real.Bottom - 8);
        int dCrop = crop.Y - (mag.Y + 34);
        bool rectOk = Math.Abs(calc.X - real.X) <= 3 && Math.Abs(calc.Y - real.Y) <= 3
                      && Math.Abs(calc.Width - real.Width) <= 3 && Math.Abs(calc.Height - real.Height) <= 4;
        bool placeOk = real.Width > 0 && Math.Abs(dLeft) <= 4 && Math.Abs(dBottom) <= 4 && Math.Abs(dCrop) <= 2;
        if (!(rectOk && placeOk)) _fail++;

        Console.WriteLine("[" + tag + "] " + ((rectOk && placeOk) ? "PASS" : "FAIL"));
        Console.WriteLine("   real(pic) = " + real.X + "," + real.Y + " " + real.Width + "x" + real.Height);
        Console.WriteLine("   rect#1    = " + calc.X + "," + calc.Y + " " + calc.Width + "x" + calc.Height +
                          "   rect#2 = " + calc2.X + "," + calc2.Y + " " + calc2.Width + "x" + calc2.Height);
        Console.WriteLine("   rect 误差 dx=" + (calc.X - real.X) + " dy=" + (calc.Y - real.Y) +
                          " dw=" + (calc.Width - real.Width) + " dh=" + (calc.Height - real.Height));
        Console.WriteLine("   mag=" + mag + "  crop=" + crop + "   贴边误差 left=" + dLeft + " bottom=" + dBottom + " cropGap=" + dCrop);
        try { File.Copy(_dump, Path.Combine(Path.GetTempPath(), "izplace_" + DateTime.Now.Ticks % 100000 + ".png"), true); } catch { }
    }

    private static string _dump = Path.Combine(Path.GetTempPath(), "izplace_shot.png");

    private static Rectangle FindMagenta()
    {
        int w = GetSystemMetrics(0), h = GetSystemMetrics(1);
        using (Bitmap shot = new Bitmap(w, h, PixelFormat.Format32bppArgb))
        {
            using (Graphics g = Graphics.FromImage(shot)) g.CopyFromScreen(0, 0, 0, 0, new Size(w, h));
            try { shot.Save(_dump, ImageFormat.Png); } catch { }
            BitmapData bd = shot.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
            try
            {
                int minX = int.MaxValue, minY = int.MaxValue, maxX = -1, maxY = -1;
                unsafe
                {
                    for (int y = 0; y < h; y++)
                    {
                        byte* row = (byte*)bd.Scan0 + y * bd.Stride;
                        for (int x = 0; x < w; x++)
                        {
                            if (row[x * 4 + 2] > 200 && row[x * 4 + 1] < 60 && row[x * 4 + 0] > 200)
                            {
                                if (x < minX) minX = x;
                                if (x > maxX) maxX = x;
                                if (y < minY) minY = y;
                                if (y > maxY) maxY = y;
                            }
                        }
                    }
                }
                if (maxX < 0) return Rectangle.Empty;
                return Rectangle.FromLTRB(minX, minY, maxX + 1, maxY + 1);
            }
            finally { shot.UnlockBits(bd); }
        }
    }

    private static void Pump()
    {
        for (int i = 0; i < 8; i++) { Application.DoEvents(); Thread.Sleep(25); }
    }

    private static Rectangle Bounds(string field)
    {
        Form f = GetField(field) as Form;
        return f == null ? Rectangle.Empty : f.Bounds;
    }

    private static object Invoke(string name, params object[] args)
    {
        MethodInfo m = _connect.GetMethod(name, BindingFlags.NonPublic | BindingFlags.Instance);
        if (m == null) { Console.WriteLine("METHOD MISSING: " + name); return null; }
        try { return m.Invoke(_conn, m.GetParameters().Length == 0 ? new object[0] : args); }
        catch (TargetInvocationException ex)
        {
            Console.WriteLine("[" + name + "] THREW: " + ex.InnerException.Message);
            return null;
        }
    }

    private static object FieldObj(object obj, string name)
    {
        FieldInfo fi = obj.GetType().GetField(name, BindingFlags.Public | BindingFlags.Instance);
        return fi == null ? null : fi.GetValue(obj);
    }

    private static string Field(object obj, string name)
    {
        object v = FieldObj(obj, name);
        return Convert.ToString(v);
    }

    private static void SetField(string name, object v)
    {
        FieldInfo fi = _connect.GetField(name, BindingFlags.NonPublic | BindingFlags.Instance);
        if (fi != null) fi.SetValue(_conn, v);
    }

    private static object GetField(string name)
    {
        FieldInfo fi = _connect.GetField(name, BindingFlags.NonPublic | BindingFlags.Instance);
        return fi == null ? null : fi.GetValue(_conn);
    }
}
