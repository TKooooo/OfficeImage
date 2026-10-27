using System;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;

/// <summary>
/// 验证二级菜单的完整链路：ShowCropMenu 开面板 -> 面板抛出选择 -> OnCropChosen 执行 -> 面板关闭。
/// 唯一没覆盖的是"鼠标按下"本身（WinForms 只认真实输入），那是用户手点的那一下。
/// </summary>
internal static class MenuProbe
{
    private static Type _connect;
    private static Type _choice;
    private static object _conn;

    [STAThread]
    private static void Main(string[] args)
    {
        Assembly asm = Assembly.LoadFrom(args[0]);
        string img = args[1];
        string outDoc = args[2];

        // WinForms 需要消息泵才有窗口消息；这里只开面板不靠定时器，简单跑一个循环即可
        _connect = asm.GetType("ImageZoomAddin.Connect");
        _choice = asm.GetType("ImageZoomAddin.CropChoice");
        _conn = Activator.CreateInstance(_connect);

        dynamic app = Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application"));
        SetField("_app", app);
        SetField("_host", "Word");
        app.DisplayAlerts = 0;
        dynamic doc = app.Documents.Add();
        dynamic pic = doc.InlineShapes.AddPicture(img);
        pic.LockAspectRatio = -1;
        pic.Width = 320f;
        pic.Select();
        Console.WriteLine("setup ok: " + R(pic.Width) + " x " + R(pic.Height));

        object hit = Invoke("DetectPicture");
        if (hit == null) { Console.WriteLine("ABORT: hit null"); return; }

        // ---- 开面板
        Invoke("ShowCropMenu", hit);
        Form menu = (Form)GetField("_menu");
        Console.WriteLine("menu: " + (menu == null ? "NULL" :
            ("visible=" + menu.Visible + " bounds=" + menu.Bounds + " size=" + menu.Size)));
        if (menu == null) { Console.WriteLine("ABORT"); return; }
        Pump();

        // ---- 选「五角星」
        Invoke("OnCropChosen", MakeChoice("Shape", "star5", "五角星"));
        Console.WriteLine("after star5: prst=" + GetPrst(doc) + "  menu=" + (GetField("_menu") == null ? "closed" : "STILL OPEN"));
        Report(doc, "after star5 ");

        // ---- 选「16:9」
        Reselect(doc);
        object h2 = Invoke("DetectPicture");
        if (h2 != null) {
            Invoke("ShowCropMenu", h2);
            Console.WriteLine("menu reopened: " + (GetField("_menu") != null));
            Invoke("OnCropChosen", MakeRatio(16, 9, "16:9"));
        }
        Report(doc, "after 16:9  ");
        Console.WriteLine("prst still = " + GetPrst(doc));

        // ---- 选「重设形状和大小」
        Reselect(doc);
        object h3 = Invoke("DetectPicture");
        if (h3 != null) Invoke("OnCropChosen", MakeOther("Reset", "重设形状和大小"));
        Console.WriteLine("after reset: prst=" + GetPrst(doc));
        Report(doc, "after reset ");

        doc.SaveAs2(outDoc, 16);
        doc.Close(0);
        app.Quit();
        Console.WriteLine("done");
    }

    // ------------------------------------------------------------ 构造面板的"选择"

    private static object MakeChoice(string kind, string prst, string label)
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

    // ------------------------------------------------------------ 反射助手

    private static void Pump()
    {
        for (int i = 0; i < 8; i++) { Application.DoEvents(); System.Threading.Thread.Sleep(25); }
    }

    private static object Invoke(string name, params object[] args)
    {
        MethodInfo m = _connect.GetMethod(name, BindingFlags.NonPublic | BindingFlags.Instance);
        if (m == null) { Console.WriteLine("METHOD MISSING: " + name); return null; }
        object[] a = m.GetParameters().Length == 0 ? new object[0] : args;
        try { return m.Invoke(_conn, a); }
        catch (TargetInvocationException ex)
        {
            Console.WriteLine("[" + name + "] THREW: " + ex.InnerException.GetType().Name +
                              " :: " + ex.InnerException.Message);
            return null;
        }
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

    private static void Reselect(dynamic doc)
    {
        try { doc.InlineShapes.Item(1).Select(); } catch { }
    }

    private static string R(object v) { return Convert.ToDouble(v).ToString("0.00"); }

    private static void Report(dynamic doc, string tag)
    {
        try
        {
            int n = Convert.ToInt32(doc.InlineShapes.Count);
            int s = Convert.ToInt32(doc.Shapes.Count);
            if (n < 1) { Console.WriteLine(tag + ": no inline (inline=" + n + " shapes=" + s + ")"); return; }
            dynamic p = doc.InlineShapes.Item(1);
            double w = Convert.ToDouble(p.Width), h = Convert.ToDouble(p.Height);
            Console.WriteLine(tag + ": " + w.ToString("0.00") + " x " + h.ToString("0.00") +
                              " ratio=" + (h > 0.01 ? (w / h).ToString("0.0000") : "n/a"));
        }
        catch (Exception ex) { Console.WriteLine(tag + ": ERR " + ex.Message); }
    }

    private static string GetPrst(dynamic doc)
    {
        try
        {
            if (Convert.ToInt32(doc.InlineShapes.Count) < 1) return "(no inline)";
            string xml = doc.InlineShapes.Item(1).Range.WordOpenXML;
            int i = xml.IndexOf("<a:prstGeom prst=\"");
            if (i < 0) return "(no prstGeom)";
            i += "<a:prstGeom prst=\"".Length;
            int j = xml.IndexOf('"', i);
            return xml.Substring(i, j - i);
        }
        catch (Exception ex) { return "ERR " + ex.Message; }
    }
}
