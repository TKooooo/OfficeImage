// 旋转功能的离线验证：
//   1) PatchTransform 的 XML 改写（角度累加 / 翻转切换 / 自闭合标签 / 属性保序）
//   2) 三个圆形按钮和旋转面板的实际渲染（放大 3 倍看图标）
// 不碰任何 Office 进程，纯反射 + 画图。
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

internal static class RotProbe
{
    private static Assembly _asm;
    private static Type _connect;
    private static Type _kind;      // RotateKind
    private static Type _btnKind;   // ButtonKind
    private static int _fail;

    private static void P(string k, object v) { Console.WriteLine("  " + k + " = " + v); }

    [STAThread]
    private static int Main(string[] args)
    {
        string dll = args.Length > 0 ? args[0] : "ImageZoomAddin.dll";
        _asm = Assembly.LoadFrom(dll);
        _connect = _asm.GetType("ImageZoomAddin.Connect");
        _kind = _asm.GetType("ImageZoomAddin.RotateKind");
        _btnKind = _asm.GetType("ImageZoomAddin.ButtonKind");
        if (_connect == null || _kind == null || _btnKind == null)
        {
            Console.WriteLine("TYPE NOT FOUND");
            return 1;
        }

        string pic = "<pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"2032000\" cy=\"1142969\"/></a:xfrm>" +
                     "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></pic:spPr>";

        // ---- 1) XML 改写
        Console.WriteLine("== PatchTransform ==");
        string once = Patch(pic, "Right90");
        P("Right90", Tag(once, "a:xfrm"));
        Check("rot=5400000", once.Contains("<a:xfrm rot=\"5400000\">"));
        Check("保留 off/ext", once.Contains("<a:off x=\"0\" y=\"0\"/>") && once.Contains("<a:ext cx=\"2032000\""));

        string back = Patch(once, "Left90");
        P("再 Left90", Tag(back, "a:xfrm"));
        Check("转回来不带 rot", !back.Contains("rot="));

        string twice = Patch(once, "Right90");
        P("再 Right90", Tag(twice, "a:xfrm"));
        Check("两次=180°", twice.Contains("rot=\"10800000\""));

        string left1 = Patch(pic, "Left90");
        P("Left90", Tag(left1, "a:xfrm"));
        Check("左转=16200000(270°)", left1.Contains("rot=\"16200000\""));

        string fh = Patch(pic, "FlipH");
        P("FlipH", Tag(fh, "a:xfrm"));
        Check("flipH=1", fh.Contains("flipH=\"1\"") && !fh.Contains("flipV"));

        string fh2 = Patch(fh, "FlipH");
        P("再 FlipH", Tag(fh2, "a:xfrm"));
        Check("翻转可逆", !fh2.Contains("flip"));

        string both = Patch(fh, "FlipV");
        P("FlipH+FlipV", Tag(both, "a:xfrm"));
        Check("两个属性共存", both.Contains("flipH=\"1\"") && both.Contains("flipV=\"1\""));

        string rotfh = Patch(Patch(pic, "Right90"), "FlipV");
        P("Right90+FlipV", Tag(rotfh, "a:xfrm"));
        Check("rot 与 flip 互不干扰", rotfh.Contains("rot=\"5400000\"") && rotfh.Contains("flipV=\"1\""));

        string self = "<a:xfrm/>";
        string selfRot = Patch(self, "Right90");
        P("自闭合 <a:xfrm/>", selfRot);
        Check("自闭合标签保留", selfRot == "<a:xfrm rot=\"5400000\"/>");

        string full = "<a:xfrm rot=\"5400000\" flipH=\"1\"><a:off x=\"1\" y=\"2\"/></a:xfrm>";
        string fullLeft = Patch(full, "Left90");
        P("多属性回退", Tag(fullLeft, "a:xfrm"));
        Check("回退后只留 flipH", fullLeft.StartsWith("<a:xfrm flipH=\"1\">"));

        string four = pic;
        for (int i = 0; i < 4; i++) four = Patch(four, "Right90");
        P("Right90 x4", Tag(four, "a:xfrm"));
        Check("转满一圈回到原样", !four.Contains("rot="));

        string none = Patch("<pic:spPr><a:prstGeom prst=\"rect\"/></pic:spPr>", "Right90");
        Check("没有 xfrm 时返回 null", none == null);

        // ---- 2) 渲染
        Console.WriteLine("== 渲染 ==");
        RenderButtons(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "buttons.png"));
        RenderMenu(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "rotate-menu.png"));

        Console.WriteLine(_fail == 0 ? "ALL PASS" : ("FAILED " + _fail));
        return _fail == 0 ? 0 : 1;
    }

    private static void Check(string name, bool ok)
    {
        if (!ok) _fail++;
        Console.WriteLine("  [" + (ok ? "PASS" : "FAIL") + "] " + name);
    }

    /// <summary>只打印某个标签本身（输出短一点）。</summary>
    private static string Tag(string xml, string tag)
    {
        if (xml == null) return "(null)";
        int i = xml.IndexOf("<" + tag);
        if (i < 0) return "(none)";
        int j = xml.IndexOf('>', i);
        return xml.Substring(i, j - i + 1);
    }

    private static string Patch(string xml, string kindName)
    {
        MethodInfo m = _connect.GetMethod("PatchTransform", BindingFlags.NonPublic | BindingFlags.Static);
        object k = Enum.Parse(_kind, kindName);
        return (string)m.Invoke(null, new object[] { xml, k });
    }

    // ------------------------------------------------------------ 渲染

    private static void RenderButtons(string path)
    {
        object[] kinds = { Enum.Parse(_btnKind, "Zoom"), Enum.Parse(_btnKind, "Crop"), Enum.Parse(_btnKind, "Rotate") };
        Type btn = _asm.GetType("ImageZoomAddin.MagnifierButton");

        using (Bitmap sheet = new Bitmap(3 * 96, 96, PixelFormat.Format32bppArgb))
        using (Graphics g = Graphics.FromImage(sheet))
        {
            g.Clear(Color.FromArgb(240, 240, 240));
            g.InterpolationMode = InterpolationMode.NearestNeighbor;
            g.PixelOffsetMode = PixelOffsetMode.Half;

            for (int i = 0; i < kinds.Length; i++)
            {
                Form f = (Form)Activator.CreateInstance(btn, new object[] { kinds[i] });
                f.StartPosition = FormStartPosition.Manual;
                f.Location = new Point(-3000, -3000);
                f.Show();
                Pump();

                using (Bitmap one = new Bitmap(f.Width, f.Height))
                {
                    f.DrawToBitmap(one, new Rectangle(0, 0, f.Width, f.Height));
                    g.DrawImage(one, new Rectangle(i * 96 + 6, 6, 84, 84));
                }
                f.Close();
            }
            sheet.Save(path, ImageFormat.Png);
        }
        Console.WriteLine("  buttons.png -> " + path);
    }

    private static void RenderMenu(string path)
    {
        Type t = _asm.GetType("ImageZoomAddin.RotateMenuForm");
        Form f = (Form)Activator.CreateInstance(t, true);
        f.StartPosition = FormStartPosition.Manual;
        f.Location = new Point(-3000, -3000);
        f.Show();
        Pump();

        Save(f, Path.GetDirectoryName(path) + "\\rotate-menu-0.png");

        FieldInfo hf = t.GetField("_hover", BindingFlags.NonPublic | BindingFlags.Instance);
        hf.SetValue(f, 2);
        f.GetType().GetMethod("Invalidate", Type.EmptyTypes).Invoke(f, null);
        Pump();
        Save(f, path);

        f.Close();
    }

    private static void Save(Form f, string path)
    {
        using (Bitmap bmp = new Bitmap(f.Width, f.Height))
        {
            f.DrawToBitmap(bmp, new Rectangle(0, 0, f.Width, f.Height));
            bmp.Save(path, ImageFormat.Png);
        }
        Console.WriteLine("  " + Path.GetFileName(path) + " -> " + path);
    }

    private static void Pump()
    {
        for (int i = 0; i < 8; i++) { Application.DoEvents(); System.Threading.Thread.Sleep(25); }
    }
}
