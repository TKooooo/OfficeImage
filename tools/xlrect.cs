// Excel 的 Window.PointsToScreenPixelsX/Y 到底怎么换算？用"品红图 + 截屏找色块"做基准，
// 分别在 100% / 50% / 150% 缩放、以及滚动工作表之后量一遍。
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;

internal static class XlRect
{
    [StructLayout(LayoutKind.Sequential)] private struct RECT { public int L, T, R, B; }

    [DllImport("user32.dll")] private static extern int GetSystemMetrics(int i);
    [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] private static extern IntPtr GetDesktopWindow();
    [DllImport("user32.dll")] private static extern IntPtr GetWindowDC(IntPtr h);
    [DllImport("user32.dll")] private static extern int ReleaseDC(IntPtr h, IntPtr dc);
    [DllImport("gdi32.dll")] private static extern int GetDeviceCaps(IntPtr dc, int index);
    [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr h, int cmd);

    private static void P(string k, object v) { Console.WriteLine("  " + k + " = " + v); }
    private static string F(dynamic v) { return Convert.ToDouble(v).ToString("0.0"); }

    private static Rectangle FindMagenta()
    {
        int w = GetSystemMetrics(0), h = GetSystemMetrics(1);
        using (Bitmap shot = new Bitmap(w, h, PixelFormat.Format32bppArgb))
        {
            using (Graphics g = Graphics.FromImage(shot)) g.CopyFromScreen(0, 0, 0, 0, new Size(w, h));
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

    private static void Report(string tag, dynamic win, dynamic shp)
    {
        double L = Convert.ToDouble(shp.Left), T = Convert.ToDouble(shp.Top);
        double W = Convert.ToDouble(shp.Width), H = Convert.ToDouble(shp.Height);
        int fx0 = win.PointsToScreenPixelsX(0), fy0 = win.PointsToScreenPixelsY(0);
        int fxL = win.PointsToScreenPixelsX(Convert.ToInt64(Math.Round(L)));
        int fyT = win.PointsToScreenPixelsY(Convert.ToInt64(Math.Round(T)));
        int fxR = win.PointsToScreenPixelsX(Convert.ToInt64(Math.Round(L + W)));
        int fyB = win.PointsToScreenPixelsY(Convert.ToInt64(Math.Round(T + H)));
        Rectangle real = FindMagenta();

        Console.WriteLine("---- " + tag + " ----");
        P("shape(pts)", "L=" + F(L) + " T=" + F(T) + " W=" + F(W) + " H=" + F(H));
        P("zoom", F(win.Zoom) + "%   scroll row/col=" + win.ScrollRow + "/" + win.ScrollColumn);
        P("f(0)", fx0 + "," + fy0);
        P("f(L),f(T)", fxL + "," + fyT);
        P("old calc", fxL + "," + fyT + " " + (fxR - fxL) + "x" + (fyB - fyT));
        P("real", real.IsEmpty ? "(not found)" : real.X + "," + real.Y + " " + real.Width + "x" + real.Height);
        if (!real.IsEmpty)
        {
            P("slopeX", ((fxR - fxL) / W).ToString("0.000") + "   trueSlopeX=" + (real.Width / W).ToString("0.000"));
            P("slopeY", ((fyB - fyT) / H).ToString("0.000") + "   trueSlopeY=" + (real.Height / H).ToString("0.000"));
            P("origin from real", "f(0) off by dx=" + (fx0 - (real.X - real.Width / W * L)) +
                                  " dy=" + (fy0 - (real.Y - real.Height / H * T)));
        }
    }

    [STAThread]
    private static int Main()
    {
        IntPtr dc = GetWindowDC(GetDesktopWindow());
        Console.WriteLine("screen=" + GetSystemMetrics(0) + "x" + GetSystemMetrics(1) +
                          "  LOGPIXELSX=" + GetDeviceCaps(dc, 88) + "  LOGPIXELSY=" + GetDeviceCaps(dc, 90));
        ReleaseDC(IntPtr.Zero, dc);

        string png = Path.Combine(Path.GetTempPath(), "izrect.png");
        using (Bitmap bmp = new Bitmap(800, 450))
        using (Graphics g = Graphics.FromImage(bmp)) { g.Clear(Color.Magenta); bmp.Save(png, ImageFormat.Png); }

        Type t = Type.GetTypeFromProgID("Excel.Application");
        dynamic xl = Activator.CreateInstance(t);
        dynamic wb = null;
        try
        {
            xl.Visible = true;
            xl.DisplayAlerts = false;
            xl.WindowState = -4137;      // xlMaximized
            wb = xl.Workbooks.Add();
            dynamic ws = wb.Worksheets.Item(1);
            dynamic win = xl.ActiveWindow;

            dynamic pic = ws.Shapes.AddPicture(png, 0, -1, 200, 200, 600, 337.5);
            pic.Name = "picprobe";
            pic.Select();

            // Excel 的窗口要等它真正画出来，PointsToScreenPixels 才会返回非 0
            for (int i = 0; i < 40; i++)
            {
                try
                {
                    if ((IntPtr)xl.Hwnd != IntPtr.Zero) ShowWindow((IntPtr)xl.Hwnd, 3); // SW_MAXIMIZE
                    SetForegroundWindow((IntPtr)xl.Hwnd);
                    if (win.PointsToScreenPixelsX(0) != 0 || win.PointsToScreenPixelsY(0) != 0) break;
                }
                catch { }
                Thread.Sleep(250);
            }
            Thread.Sleep(600);
            pic.Select();
            Thread.Sleep(400);

            Console.WriteLine("hwnd=" + xl.Hwnd);
            RECT wr; GetWindowRect((IntPtr)xl.Hwnd, out wr);
            Console.WriteLine("excel window rect = " + wr.L + "," + wr.T + " - " + wr.R + "," + wr.B);

            Report("zoom100", win, pic);
            win.Zoom = 50; Thread.Sleep(800); Report("zoom50", win, pic);
            win.Zoom = 150; Thread.Sleep(800); Report("zoom150", win, pic);
            win.Zoom = 100; Thread.Sleep(800);
            win.ScrollRow = 6; win.ScrollColumn = 3; Thread.Sleep(800); Report("scroll6/3", win, pic);
        }
        catch (Exception ex) { Console.WriteLine("ERR " + ex); }
        finally
        {
            try { if (wb != null) wb.Close(false); } catch { }
            try { xl.Quit(); } catch { }
            try { Marshal.ReleaseComObject(xl); } catch { }
        }
        return 0;
    }
}
