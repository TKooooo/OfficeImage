// PowerPoint 坐标换算验证（第二版）。
//
// 截图找色块在 PPT 上不可靠（屏幕上别的品红元素会混进来），改用 PowerPoint 自己的
// DocumentWindow.RangeFromPoint(x, y)：给屏幕像素坐标，返回该处的 Shape。
// 从图片中心点向四个方向逐像素外扩，能精确量出图片在屏幕上的矩形，再跟公式算的比。
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;

internal static class PptProbe
{
    [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr h, int cmd);
    [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] private static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint flags);
    [DllImport("user32.dll")] private static extern bool BringWindowToTop(IntPtr h);
    [DllImport("user32.dll")] private static extern IntPtr GetDesktopWindow();
    [DllImport("user32.dll")] private static extern IntPtr GetWindowDC(IntPtr h);
    [DllImport("user32.dll")] private static extern int ReleaseDC(IntPtr h, IntPtr dc);
    [DllImport("gdi32.dll")] private static extern int GetDeviceCaps(IntPtr dc, int index);

    private static void P(string k, object v) { Console.WriteLine("  " + k + " = " + v); }
    private static string F(dynamic v) { try { return Convert.ToDouble(v).ToString("0.00"); } catch { return "?"; } }

    private static IntPtr _hwnd = IntPtr.Zero;
    private static dynamic _win;
    private static string _name = "";

    private static void Front()
    {
        if (_hwnd == IntPtr.Zero) return;
        SetWindowPos(_hwnd, new IntPtr(-1), 0, 0, 0, 0, 0x0043);
        BringWindowToTop(_hwnd);
        SetForegroundWindow(_hwnd);
        Thread.Sleep(250);
    }

    /// <summary>屏幕点 (x,y) 落在我们那张图片上吗？</summary>
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

    /// <summary>从中心点向四个方向逐像素外扩，量出图片在屏幕上的矩形。</summary>
    private static Rectangle Measure(int cx, int cy)
    {
        if (!Hit(cx, cy)) return Rectangle.Empty;

        int l = cx;
        while (l > 1 && Hit(l - 1, cy)) l--;
        int r = cx;
        while (r < 3999 && Hit(r + 1, cy)) r++;
        int t = cy;
        while (t > 1 && Hit(cx, t - 1)) t--;
        int b = cy;
        while (b < 2999 && Hit(cx, b + 1)) b++;

        return Rectangle.FromLTRB(l, t, r + 1, b + 1);
    }

    [STAThread]
    private static int Main()
    {
        IntPtr dc = GetWindowDC(GetDesktopWindow());
        Console.WriteLine("screen=" + System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width + "x" +
                          System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height +
                          "  LOGPIXELSX=" + GetDeviceCaps(dc, 88) + "  LOGPIXELSY=" + GetDeviceCaps(dc, 90));
        ReleaseDC(IntPtr.Zero, dc);

        string png = Path.Combine(Path.GetTempPath(), "izppt.png");
        using (Bitmap bmp = new Bitmap(800, 450))
        using (Graphics g = Graphics.FromImage(bmp)) { g.Clear(Color.Magenta); bmp.Save(png, ImageFormat.Png); }

        Type t = Type.GetTypeFromProgID("PowerPoint.Application");
        if (t == null) { Console.WriteLine("PowerPoint 未安装"); return 1; }

        dynamic app = Activator.CreateInstance(t);
        dynamic prs = null;
        try
        {
            app.Visible = true;
            prs = app.Presentations.Add();
            dynamic slide = prs.Slides.Add(1, 12);
            _win = app.ActiveWindow;

            dynamic pic = slide.Shapes.AddPicture(png, 0, -1, 200, 200, 600, 337.5);
            pic.Name = "picprobe";
            _name = "picprobe";

            for (int i = 0; i < 20; i++)
            {
                foreach (System.Diagnostics.Process pr in System.Diagnostics.Process.GetProcessesByName("POWERPNT"))
                    if (pr.MainWindowHandle != IntPtr.Zero) { _hwnd = pr.MainWindowHandle; break; }
                if (_hwnd != IntPtr.Zero) break;
                Thread.Sleep(300);
            }
            Console.WriteLine("hwnd=" + _hwnd);
            try { app.WindowState = 3; } catch { }
            ShowWindow(_hwnd, 3);
            Front();

            pic.Select();
            Thread.Sleep(600);
            pic.Rotation = 0;

            Console.WriteLine("slide=" + F(prs.PageSetup.SlideWidth) + "x" + F(prs.PageSetup.SlideHeight) +
                              "  shape=" + F(pic.Left) + "," + F(pic.Top) + " " + F(pic.Width) + "x" + F(pic.Height));

            foreach (double z in new double[] { 100, 50, 150, 66, 200 })
            {
                try { _win.View.Zoom = (float)z; } catch { }
                Thread.Sleep(700);
                Front();
                Report("zoom" + z, pic);
            }

            try { _win.View.Zoom = 100; } catch { }
            Thread.Sleep(600);
            Front();

            pic.Rotation = 90;
            Thread.Sleep(500);
            Front();
            Report("rot90", pic);
            pic.Rotation = 0;

            // 滚动到下一张再回来，看 anchor 是否仍然线性
            try
            {
                prs.Slides.Add(2, 12);
                _win.View.GotoSlide(2);
                Thread.Sleep(400);
                _win.View.GotoSlide(1);
                Thread.Sleep(500);
                Front();
                Report("back-to-1", pic);
            }
            catch (Exception ex) { P("goto", ex.Message); }
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

    private static void Report(string tag, dynamic pic)
    {
        Console.WriteLine("---- " + tag + " ----");

        double L = Convert.ToDouble(pic.Left), T = Convert.ToDouble(pic.Top);
        double W = Convert.ToDouble(pic.Width), H = Convert.ToDouble(pic.Height);

        double zoom = 1.0;
        try { zoom = Convert.ToDouble(_win.View.Zoom) / 100.0; } catch { }

        int fx0, fy0;
        try
        {
            fx0 = Convert.ToInt32(_win.PointsToScreenPixelsX(0f));
            fy0 = Convert.ToInt32(_win.PointsToScreenPixelsY(0f));
        }
        catch (Exception ex) { Console.WriteLine("  f() ERR " + ex.Message); return; }

        double k = 1.3333 * zoom;
        int x1 = fx0 + (int)Math.Round(L * k);
        int y1 = fy0 + (int)Math.Round(T * k);
        int rw = (int)Math.Round(W * k);
        int rh = (int)Math.Round(H * k);

        P("zoom", (zoom * 100).ToString("0") + "%   k=" + k.ToString("0.0000"));
        P("f(0)", fx0 + "," + fy0);
        P("calc", x1 + "," + y1 + " " + rw + "x" + rh);

        Rectangle real = Measure(x1 + rw / 2, y1 + rh / 2);
        P("real(RangeFromPoint)", real.IsEmpty ? "(miss)" : real.X + "," + real.Y + " " + real.Width + "x" + real.Height);
        if (!real.IsEmpty)
        {
            P("delta", "dx=" + (x1 - real.X) + " dy=" + (y1 - real.Y) +
                        "  dw=" + (rw - real.Width) + " dh=" + (rh - real.Height));
            P("true k", (real.Width / W).ToString("0.0000") + " / " + (real.Height / H).ToString("0.0000"));
        }
    }
}
