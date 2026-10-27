using System;
using System.Reflection;

/// <summary>
/// 端到端验证：把真实的 Word.Application 塞进插件实例，直接调内部方法，
/// 验证「裁剪成形状 / 按比例裁剪 / 重设」这三个操作对真实文档做了什么。
/// 不模拟鼠标点击（WinForms 只认真实输入），只验证操作本身。
/// </summary>
internal static class ActionProbe
{
    private static Type _connect;
    private static object _conn;

    private static void Main(string[] args)
    {
        string dll = args[0];
        string img = args[1];
        string outDoc = args[2];

        Assembly asm = Assembly.LoadFrom(dll);
        _connect = asm.GetType("ImageZoomAddin.Connect");
        _conn = Activator.CreateInstance(_connect);
        Console.WriteLine("connect created");

        dynamic app = Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application"));
        SetField("_app", app);
        SetField("_host", "Word");
        Console.WriteLine("word attached: " + app.Version);

        app.DisplayAlerts = 0;
        dynamic doc = app.Documents.Add();
        dynamic pic = doc.InlineShapes.AddPicture(img);
        pic.LockAspectRatio = -1;
        pic.Width = 320f;
        pic.Select();
        Console.WriteLine("inserted: " + R(pic.Width) + " x " + R(pic.Height));

        object hit = Invoke("DetectPicture");
        Console.WriteLine("hit: " + (hit == null ? "NULL" : ToStr(hit, "Key")));
        if (hit == null) { Console.WriteLine("ABORT: DetectPicture returned null"); return; }

        // ---- 1. 裁剪成椭圆
        Invoke("ApplyShapeCrop", hit, "ellipse", "椭圆");
        Console.WriteLine("prst after shape = " + GetPrst(doc));
        Report(doc, "after shape      ");
        Console.WriteLine("selection after shape = " + app.Selection.Type + " (7=图片)");

        // 选中状态丢了就重新选中（模拟用户再点一下图片）
        Reselect(doc);
        Console.WriteLine("selection after reselect = " + app.Selection.Type);

        // ---- 2. 1:1
        Reselect(doc);
        object h2 = Invoke("DetectPicture");
        Console.WriteLine("hit2: " + (h2 == null ? "NULL" : "ok"));
        if (h2 != null) Invoke("ApplyRatioCrop", h2, 1.0, 1.0, "1:1");
        Report(doc, "after ratio 1:1  ");

        // ---- 3. 16:9
        Reselect(doc);
        object h3 = Invoke("DetectPicture");
        Console.WriteLine("hit3: " + (h3 == null ? "NULL" : "ok"));
        if (h3 != null) Invoke("ApplyRatioCrop", h3, 16.0, 9.0, "16:9");
        Report(doc, "after ratio 16:9 ");
        Console.WriteLine("prst still = " + GetPrst(doc));

        // ---- 4. 16:9 -> 9:16（需要从宽裁改成高裁，验证换向）
        Reselect(doc);
        object h3b = Invoke("DetectPicture");
        Console.WriteLine("hit3b: " + (h3b == null ? "NULL" : "ok"));
        if (h3b != null) Invoke("ApplyRatioCrop", h3b, 9.0, 16.0, "9:16");
        Report(doc, "after ratio 9:16 ");

        // ---- 5. 重设形状和大小
        Reselect(doc);
        object h4 = Invoke("DetectPicture");
        Console.WriteLine("hit4: " + (h4 == null ? "NULL" : "ok"));
        if (h4 != null) Invoke("ResetPictureCmd", h4);
        Console.WriteLine("prst after reset = " + GetPrst(doc));
        Report(doc, "after reset      ");

        doc.SaveAs2(outDoc, 16);
        doc.Close(0);
        app.Quit();
        Console.WriteLine("saved " + outDoc);
    }

    private static void Reselect(dynamic doc)
    {
        try
        {
            int n = Convert.ToInt32(doc.InlineShapes.Count);
            if (n >= 1) { doc.InlineShapes.Item(1).Select(); return; }
            int s = Convert.ToInt32(doc.Shapes.Count);
            if (s >= 1) doc.Shapes.Item(1).Select();
        }
        catch (Exception ex) { Console.WriteLine("Reselect failed: " + ex.Message); }
    }

    private static string R(object v) { return Convert.ToDouble(v).ToString("0.00"); }

    private static object Invoke(string name, params object[] args)
    {
        MethodInfo m = _connect.GetMethod(name, BindingFlags.NonPublic | BindingFlags.Instance);
        if (m == null) { Console.WriteLine("METHOD MISSING: " + name); return null; }
        object[] a = args;
        if (m.GetParameters().Length == 0) a = new object[0];
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
        if (fi == null) { Console.WriteLine("FIELD MISSING: " + name); return; }
        fi.SetValue(_conn, v);
    }

    private static string ToStr(object target, string field)
    {
        FieldInfo fi = target.GetType().GetField(field);
        if (fi == null) return "(no field)";
        object v = fi.GetValue(target);
        return v == null ? "(null)" : v.ToString();
    }

    private static void Report(dynamic doc, string tag)
    {
        try
        {
            int n = Convert.ToInt32(doc.InlineShapes.Count);
            int s = Convert.ToInt32(doc.Shapes.Count);
            if (n < 1) { Console.WriteLine(tag + ": no INLINE shape (inline=" + n + " shapes=" + s + ")"); return; }
            dynamic p = doc.InlineShapes.Item(1);
            double w = Convert.ToDouble(p.Width), h = Convert.ToDouble(p.Height);
            Console.WriteLine(tag + ": " + w.ToString("0.00") + " x " + h.ToString("0.00") +
                              "  ratio=" + (h > 0.01 ? (w / h).ToString("0.0000") : "n/a") +
                              "  inline=" + n + " shapes=" + s);
        }
        catch (Exception ex) { Console.WriteLine(tag + ": ERR " + ex.Message); }
    }

    private static string GetPrst(dynamic doc)
    {
        try
        {
            if (Convert.ToInt32(doc.InlineShapes.Count) < 1) return "(no inline shape)";
            dynamic p = doc.InlineShapes.Item(1);
            string xml = p.Range.WordOpenXML;
            int i = xml.IndexOf("<a:prstGeom prst=\"");
            if (i < 0) return "(no prstGeom)";
            i += "<a:prstGeom prst=\"".Length;
            int j = xml.IndexOf('"', i);
            return xml.Substring(i, j - i);
        }
        catch (Exception ex) { return "ERR " + ex.Message; }
    }
}
