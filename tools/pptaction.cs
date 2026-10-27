// PowerPoint 端到端动作验证：把真实的 PowerPoint.Application 塞进加载项实例，
// 直接调内部方法，看「取原图 / 裁剪为形状 / 按比例裁剪 / 旋转 / 重设」对真实演示文稿做了什么。
// 不模拟鼠标点击（WinForms 只认真实输入），只验证操作本身。
using System;
using System.Reflection;

internal static class PptAction
{
    private static Type _connect;
    private static object _conn;
    private static Assembly _asm;

    private static void Main(string[] args)
    {
        string dll = args[0];
        string img = args[1];
        string outFile = args[2];

        _asm = Assembly.LoadFrom(dll);
        _connect = _asm.GetType("ImageZoomAddin.Connect");
        _conn = Activator.CreateInstance(_connect);
        Console.WriteLine("connect created");

        dynamic app = Activator.CreateInstance(Type.GetTypeFromProgID("PowerPoint.Application"));
        SetField("_app", app);
        SetField("_host", "PowerPoint");
        Console.WriteLine("ppt attached: " + app.Version);

        dynamic prs = app.Presentations.Add();
        dynamic slide = prs.Slides.Add(1, 12);
        dynamic pic = slide.Shapes.AddPicture(img, 0, -1, 200, 200, 600, 337.5);
        pic.Name = "actpic";
        pic.Select();
        System.Threading.Thread.Sleep(600);
        Console.WriteLine("inserted: " + R(pic.Width) + " x " + R(pic.Height) +
                          "  id=" + pic.Id + "  ast=" + pic.AutoShapeType);

        // ---------------- 1. 选中识别 ----------------
        Console.WriteLine("---- 1. DetectPicture ----");
        object hit = Invoke("DetectPicture");
        if (hit == null) { Console.WriteLine("ABORT: null"); return; }
        Console.WriteLine("  Key        = " + ToStr(hit, "Key"));
        Console.WriteLine("  Host       = " + ToStr(hit, "Host"));
        Console.WriteLine("  Name       = " + ToStr(hit, "Name"));
        Console.WriteLine("  ShapeId    = " + ToStr(hit, "ShapeId"));
        Console.WriteLine("  SlideIndex = " + ToStr(hit, "SlideIndex"));
        object rect = GetField(hit, "ScreenRect");
        Console.WriteLine("  ScreenRect = " + (rect == null ? "(null)" : ((System.Drawing.Rectangle)rect).ToString()));

        // ---------------- 2. 取原图 ----------------
        Console.WriteLine("---- 2. Extract ----");
        object payload = Invoke("Extract", hit);
        if (payload == null) { Console.WriteLine("  !! payload null —— 拿不到原图，放大镜就废了"); }
        else
        {
            object bytes = GetField(payload, "Bytes");
            byte[] b = bytes as byte[];
            Console.WriteLine("  Source = " + ToStr(payload, "Source"));
            Console.WriteLine("  Bytes  = " + (b == null ? "null" : b.Length + " bytes"));
            if (b != null && b.Length > 16)
                Console.WriteLine("  Magic  = " + b[0].ToString("X2") + " " + b[1].ToString("X2") + " " + b[2].ToString("X2"));
            // 存下来给人工看
            if (b != null) System.IO.File.WriteAllBytes(outFile + ".png", b);
        }

        // ---------------- 3. 裁剪成椭圆 ----------------
        Console.WriteLine("---- 3. 裁剪为形状（椭圆）----");
        Invoke("ApplyPowerPointShapeCrop", hit, "ellipse", "椭圆");
        Console.WriteLine("  ast = " + pic.AutoShapeType + " (9=ellipse)   type=" + pic.Type + " (13=图片)");

        // ---------------- 4. 形状之上再按 16:9 ----------------
        Console.WriteLine("---- 4. 椭圆上再裁 16:9 ----");
        object h2 = Invoke("DetectPicture");
        if (h2 != null) Invoke("ApplyRatioCrop", h2, 16.0, 9.0, "16:9");
        Report(pic, "  16:9");

        Console.WriteLine("---- 5. 再裁 9:16 ----");
        object h3 = Invoke("DetectPicture");
        if (h3 != null) Invoke("ApplyRatioCrop", h3, 9.0, 16.0, "9:16");
        Report(pic, "  9:16");

        // ---------------- 6. 旋转 / 翻转 ----------------
        Console.WriteLine("---- 6. 旋转 ----");
        object h4 = Invoke("DetectPicture");
        object right90 = Enum.Parse(_asm.GetType("ImageZoomAddin.RotateKind"), "Right90");
        object left90 = Enum.Parse(_asm.GetType("ImageZoomAddin.RotateKind"), "Left90");
        object flipH = Enum.Parse(_asm.GetType("ImageZoomAddin.RotateKind"), "FlipH");
        if (h4 != null)
        {
            Invoke("ApplyShapeRotate", h4, right90);
            Console.WriteLine("  Right90 -> rotation=" + R(pic.Rotation));
            Invoke("ApplyShapeRotate", h4, right90);
            Console.WriteLine("  Right90 -> rotation=" + R(pic.Rotation));
            Invoke("ApplyShapeRotate", h4, left90);
            Console.WriteLine("  Left90  -> rotation=" + R(pic.Rotation));
            Invoke("ApplyShapeRotate", h4, flipH);
            Console.WriteLine("  FlipH   -> rotation=" + R(pic.Rotation));
            Invoke("ApplyShapeRotate", h4, flipH);
            Console.WriteLine("  FlipH   -> rotation=" + R(pic.Rotation));
            Invoke("ApplyShapeRotate", h4, left90);
            Console.WriteLine("  Left90  -> rotation=" + R(pic.Rotation));
        }

        // ---------------- 7. 重设 ----------------
        Console.WriteLine("---- 7. 重设形状和大小 ----");
        object h5 = Invoke("DetectPicture");
        if (h5 != null) Invoke("ResetPictureCmd", h5);
        Console.WriteLine("  ast = " + pic.AutoShapeType + " (1=矩形)   " + R(pic.Width) + " x " + R(pic.Height) +
                          "  crop=" + R(pic.PictureFormat.CropLeft) + "/" + R(pic.PictureFormat.CropRight) + "/" +
                          R(pic.PictureFormat.CropTop) + "/" + R(pic.PictureFormat.CropBottom));

        // ---------------- 8. 自由裁剪（Ribbon 命令能不能执行） ----------------
        // 不走 ExecuteNativeCrop：它失败时会弹 MessageBox，自动化会卡住。
        // 这里直接试 ExecuteMso，只看抛不抛异常。
        Console.WriteLine("---- 8. 自由裁剪 PictureCrop ----");
        try
        {
            pic.Select();
            app.ActiveWindow.Activate();
            object en = app.CommandBars.GetEnabledMso("PictureCrop");
            Console.WriteLine("  GetEnabledMso = " + en);
            app.CommandBars.ExecuteMso("PictureCrop");
            Console.WriteLine("  ExecuteMso(PictureCrop) OK —— 进入了裁剪状态");
            System.Threading.Thread.Sleep(400);
            // 退出裁剪状态
            app.CommandBars.ExecuteMso("PictureResetAndSize");
            Console.WriteLine("  已退出裁剪状态");
        }
        catch (Exception ex) { Console.WriteLine("  ExecuteMso FAIL: " + ex.Message); }

        // ---------------- 9. 存一份给人看 ----------------
        string abs = System.IO.Path.GetFullPath(outFile);
        prs.SaveAs(abs);
        Console.WriteLine("saved " + abs);
        prs.Close();
        app.Quit();
    }

    private static void Report(dynamic pic, string tag)
    {
        double w = Convert.ToDouble(pic.Width), h = Convert.ToDouble(pic.Height);
        Console.WriteLine(tag + ": " + w.ToString("0.00") + " x " + h.ToString("0.00") +
                          "  ratio=" + (h > 0.01 ? (w / h).ToString("0.0000") : "n/a") +
                          "  crop=" + R(pic.PictureFormat.CropLeft) + "/" + R(pic.PictureFormat.CropRight) + "/" +
                          R(pic.PictureFormat.CropTop) + "/" + R(pic.PictureFormat.CropBottom));
    }

    private static string R(object v) { return Convert.ToDouble(v).ToString("0.00"); }

    private static object Invoke(string name, params object[] args)
    {
        MethodInfo m = _connect.GetMethod(name, BindingFlags.NonPublic | BindingFlags.Instance);
        if (m == null) { Console.WriteLine("  METHOD MISSING: " + name); return null; }
        object[] a = args;
        if (m.GetParameters().Length == 0) a = new object[0];
        try { return m.Invoke(_conn, a); }
        catch (TargetInvocationException ex)
        {
            Console.WriteLine("  [" + name + "] THREW: " + ex.InnerException.GetType().Name +
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

    private static object GetField(object target, string field)
    {
        FieldInfo fi = target.GetType().GetField(field);
        return fi == null ? null : fi.GetValue(target);
    }

    private static string ToStr(object target, string field)
    {
        object v = GetField(target, field);
        return v == null ? "(null)" : v.ToString();
    }
}
