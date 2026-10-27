// Excel 端到端干跑（不点鼠标，直接反射调加载项的私有方法）
//   1) ShowMagnifier 在图片缩放后是否重新定位（对应需求 1）
//   2) 裁剪二级菜单在 Excel 下的四类动作（对应需求 2）
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

internal static class XlMenuProbe
{
    private static Type _connect;
    private static Type _choice;
    private static object _conn;
    private static dynamic _pic;

    [STAThread]
    private static int Main(string[] args)
    {
        Assembly asm = Assembly.LoadFrom(args[0]);
        string img = Path.Combine(Path.GetTempPath(), "izmenu_src.png");
        using (Bitmap bmp = new Bitmap(900, 600))
        using (Graphics g = Graphics.FromImage(bmp))
        {
            g.Clear(Color.CornflowerBlue);
            g.DrawRectangle(Pens.White, 8, 8, 884, 584);
            bmp.Save(img, ImageFormat.Png);
        }

        _connect = asm.GetType("ImageZoomAddin.Connect");
        _choice = asm.GetType("ImageZoomAddin.CropChoice");
        _conn = Activator.CreateInstance(_connect);

        dynamic xl = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
        SetField("_app", (object)xl);
        SetField("_host", "Excel");
        xl.DisplayAlerts = false;
        dynamic wb = xl.Workbooks.Add();
        dynamic ws = wb.Worksheets.Item(1);

        _pic = ws.Shapes.AddPicture(img, 0, -1, 120, 120, 400, 300);
        _pic.Name = "menu_img";
        _pic.Select();
        Console.WriteLine("setup: " + _pic.Width + "x" + _pic.Height + " selected=" + xl.Selection.ShapeRange.Item(1).Name);

        // ---------------- 需求 1：缩放后按钮是否跟随
        object hit = Invoke("DetectPicture");
        if (hit == null) { Console.WriteLine("ABORT: hit=null"); Cleanup(xl, wb); return 1; }
        Console.WriteLine("hit: " + Field(hit, "Key") + " rect=" + Field(hit, "ScreenRect"));

        Invoke("ShowMagnifier", hit);
        Pump();
        Rectangle m1 = Bounds("_magnifier"), c1 = Bounds("_crop");
        Console.WriteLine("after show   : mag=" + m1 + " crop=" + c1);

        _pic.Width = 220;              // 缩小图片
        _pic.Height = 160;
        object hit2 = Invoke("DetectPicture");
        Invoke("ShowMagnifier", hit2);
        Pump();
        Rectangle m2 = Bounds("_magnifier"), c2 = Bounds("_crop");
        Console.WriteLine("after resize : mag=" + m2 + " crop=" + c2);
        Console.WriteLine("followed     : " + (m1.Location != m2.Location ? "YES" : "NO"));

        // ---------------- 需求 2：裁剪菜单四类动作
        // (a) 裁剪为形状：五角星
        object h = Invoke("DetectPicture");
        Invoke("ShowCropMenu", h);
        Form menu = (Form)GetField("_menu");
        Console.WriteLine("menu open    : " + (menu != null) + (menu == null ? "" : " " + menu.Bounds));
        Pump();
        Invoke("OnCropChosen", Make("Shape", "star5", "五角星"));
        Console.WriteLine("after star5  : " + Report(ws));
        object h2 = Invoke("DetectPicture");
        Console.WriteLine("hit after    : " + (h2 == null ? "NULL" : Field(h2, "Key") + " shapeImage=" + Field(h2, "IsShapeImage")));

        // (b) 换成椭圆，验证"形状套形状"能正确替换
        if (h2 != null)
        {
            Invoke("ShowCropMenu", h2);
            Invoke("OnCropChosen", Make("Shape", "ellipse", "椭圆"));
            Console.WriteLine("after ellipse: " + Report(ws));
        }

        // (c) 按比例 16:9 —— 应先把形状还原成图片再裁
        object h3 = Invoke("DetectPicture");
        if (h3 != null)
        {
            Invoke("ShowCropMenu", h3);
            Invoke("OnCropChosen", MakeRatio(16, 9, "16:9"));
            Console.WriteLine("after 16:9   : " + Report(ws));
        }

        // (d) 重设形状和大小
        object h4 = Invoke("DetectPicture");
        if (h4 != null)
        {
            Invoke("ShowCropMenu", h4);
            Invoke("OnCropChosen", MakeOther("Reset", "重设形状和大小"));
            Console.WriteLine("after reset  : " + Report(ws));
        }

        Console.WriteLine("visible orig : " + _pic.Visible);
        Cleanup(xl, wb);
        Console.WriteLine("DONE");
        return 0;
    }

    private static void Cleanup(dynamic xl, dynamic wb)
    {
        try { wb.Close(false); } catch { }
        try { xl.Quit(); } catch { }
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }

    private static string Report(dynamic ws)
    {
        try
        {
            int n = Convert.ToInt32(ws.Shapes.Count);
            string s = "shapes=" + n;
            for (int i = 1; i <= n; i++)
            {
                dynamic sp = ws.Shapes.Item(i);
                s += " | " + sp.Name + " type=" + sp.Type + " vis=" + sp.Visible +
                     " " + Convert.ToDouble(sp.Width).ToString("0.0") + "x" + Convert.ToDouble(sp.Height).ToString("0.0");
                if (Convert.ToInt32(sp.Type) == 1) s += " fill=" + sp.Fill.Type + " alt=" + sp.AlternativeText;
            }
            return s;
        }
        catch (Exception ex) { return "ERR " + ex.Message; }
    }

    private static object Make(string kind, string prst, string label)
    {
        MethodInfo m = _choice.GetMethod("Shape", BindingFlags.Public | BindingFlags.Static);
        return m.Invoke(null, new object[] { prst, label });
    }

    private static object MakeRatio(double rw, double rh, string label)
    {
        MethodInfo m = _choice.GetMethod("Ratio", BindingFlags.Public | BindingFlags.Static);
        return m.Invoke(null, new object[] { rw, rh, label });
    }

    private static object MakeOther(string kind, string label)
    {
        MethodInfo m = _choice.GetMethod("Other", BindingFlags.Public | BindingFlags.Static);
        Type kt = m.GetParameters()[0].ParameterType;
        return m.Invoke(null, new object[] { Enum.Parse(kt, kind), label });
    }

    private static void Pump()
    {
        for (int i = 0; i < 10; i++) { Application.DoEvents(); System.Threading.Thread.Sleep(25); }
    }

    private static Rectangle Bounds(string field)
    {
        object o = GetField(field);
        Form f = o as Form;
        return f == null ? Rectangle.Empty : f.Bounds;
    }

    private static object Invoke(string name, params object[] args)
    {
        MethodInfo m = _connect.GetMethod(name, BindingFlags.NonPublic | BindingFlags.Instance);
        if (m == null) { Console.WriteLine("METHOD MISSING: " + name); return null; }
        object[] a = m.GetParameters().Length == 0 ? new object[0] : args;
        try { return m.Invoke(_conn, a); }
        catch (TargetInvocationException ex)
        {
            Console.WriteLine("[" + name + "] THREW: " + ex.InnerException.GetType().Name + " :: " + ex.InnerException.Message);
            return null;
        }
    }

    private static string Field(object obj, string name)
    {
        FieldInfo fi = obj.GetType().GetField(name, BindingFlags.Public | BindingFlags.Instance);
        if (fi == null) return "?";
        object v = fi.GetValue(obj);
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
