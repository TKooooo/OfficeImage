using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;

/// <summary>
/// 在真实桌面上把一个按钮显示出来，确认分层窗口（UpdateLayeredWindow）确实把内容画上去了：
/// 中心应该是白的（按钮本体），四角应该是桌面原色（圆角外透明，说明 alpha 生效）。
/// 只看 30x30 的一块，1 秒内关掉。
/// </summary>
class LayeredProbe
{
    [STAThread]
    static void Main(string[] args)
    {
        Assembly asm = Assembly.LoadFrom(args[0]);
        Type btnType = asm.GetType("ImageZoomAddin.MagnifierButton");
        Type kindType = asm.GetType("ImageZoomAddin.ButtonKind");

        Rectangle work = Screen.PrimaryScreen.WorkingArea;
        Point at = new Point(work.Right - 180, work.Bottom - 120);

        Form f = (Form)Activator.CreateInstance(btnType, new object[] { Enum.Parse(kindType, "Zoom") });
        f.StartPosition = FormStartPosition.Manual;
        f.Location = at;
        f.Show();

        for (int i = 0; i < 12; i++) { Application.DoEvents(); Thread.Sleep(30); }

        FieldInfo lf = btnType.GetField("_layeredOk", BindingFlags.NonPublic | BindingFlags.Instance);
        Console.WriteLine("_layeredOk = " + lf.GetValue(f));

        using (Bitmap shot = new Bitmap(30, 30, PixelFormat.Format32bppArgb))
        {
            using (Graphics g = Graphics.FromImage(shot))
                g.CopyFromScreen(at.X, at.Y, 0, 0, new Size(30, 30));
            shot.Save(args[1], ImageFormat.Png);
            Console.WriteLine("center(15,15) = " + shot.GetPixel(15, 15));
            Console.WriteLine("corner(0,0)   = " + shot.GetPixel(0, 0));
            Console.WriteLine("corner(29,29) = " + shot.GetPixel(29, 29));
            Console.WriteLine("edge(0,15)    = " + shot.GetPixel(0, 15));
        }
        f.Close();
    }
}
