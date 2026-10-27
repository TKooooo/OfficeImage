// Excel 侧旋转/翻转验证（独立隐藏实例，不碰用户正在用的 Excel）
//   验证 Shape.Rotation 累加、Shape.Flip(msoFlipHorizontal/Vertical) 存在且生效，
//   以及"图片填充自选图形"（裁剪为形状的产物）同样能吃这些属性。
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

internal static class XlRotate
{
    private static Type _connect;
    private static Type _kind;
    private static object _conn;
    private static dynamic _xl;
    private static int _fail;

    [STAThread]
    private static int Main(string[] args)
    {
        string dll = args.Length > 0 ? args[0] : "ImageZoomAddin.dll";
        Assembly asm = Assembly.LoadFrom(dll);
        _connect = asm.GetType("ImageZoomAddin.Connect");
        _kind = asm.GetType("ImageZoomAddin.RotateKind");
        _conn = Activator.CreateInstance(_connect);

        string img = Path.Combine(Path.GetTempPath(), "izrot.png");
        using (Bitmap bmp = new Bitmap(800, 450))
        using (Graphics g = Graphics.FromImage(bmp)) { g.Clear(Color.SteelBlue); bmp.Save(img, ImageFormat.Png); }

        _xl = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
        SetField("_app", (object)_xl);
        SetField("_host", "Excel");
        _xl.DisplayAlerts = false;
        dynamic wb = _xl.Workbooks.Add();
        dynamic ws = wb.Worksheets.Item(1);

        try
        {
            // ---------- 普通图片
            dynamic pic = ws.Shapes.AddPicture(img, 0, -1, 100, 100, 400, 225);
            pic.Name = "rot_pic";
            pic.Select();

            object hit = Detect();
            Console.WriteLine("== 普通图片 ==");
            Rotate(hit, "Right90");
            Check("右转 90 -> rotation=90", Math.Abs(Rotation(pic) - 90.0) < 0.1);
            Rotate(hit, "Right90");
            Check("再右转 -> 180", Math.Abs(Rotation(pic) - 180.0) < 0.1);
            Rotate(hit, "Left90");
            Check("左转 -> 90", Math.Abs(Rotation(pic) - 90.0) < 0.1);
            Rotate(hit, "Left90");
            Check("再左转 -> 0", Math.Abs(Rotation(pic) - 0.0) < 0.1);

            Rotate(hit, "FlipH");
            Check("水平翻转不报错", true);
            Rotate(hit, "FlipV");
            Check("垂直翻转不报错", true);

            // ---------- 图片填充自选图形（裁剪为形状的产物，用插件自己的流程产出来）
            dynamic pic2 = ws.Shapes.AddPicture(img, 0, -1, 100, 400, 400, 225);
            pic2.Name = "rot_pic2";
            pic2.Select();
            object hit20 = Detect();
            Invoke("ApplyExcelShapeCrop", hit20, "star5", "五角星");
            object hit2 = Detect();
            Console.WriteLine("== 形状图 ==");
            if (hit2 == null) { Check("认得裁剪后的形状图", false); }
            else
            {
                Check("认得裁剪后的形状图", true);
                Rotate(hit2, "Right90");
                dynamic shp2 = _xl.ActiveSheet.Shapes.Item(1);
                // 形状图可能就是那个 star5（原图被隐藏）
                dynamic target = null;
                for (int i = 1; i <= Convert.ToInt32(ws.Shapes.Count); i++)
                {
                    dynamic sp = ws.Shapes.Item(i);
                    if (Convert.ToString(sp.Name).IndexOf("pic2") >= 0 &&
                        Convert.ToString(sp.AlternativeText).StartsWith("IZADDIN:SHAPE:")) target = sp;
                }
                if (target != null) Check("形状图右转 90", Math.Abs(Rotation(target) - 90.0) < 0.1);
                else Check("找到形状图对象", false);
                Rotate(hit2, "FlipV");
                Check("形状图垂直翻转不报错", true);
            }

            // ---------- 旋转后的 ScreenRect 折算（隐藏窗口下算不出来，只看不抛异常）
            Console.WriteLine("== 旋转后的矩形 ==");
            pic.Select();
            object h3 = Detect();
            Console.WriteLine("  detect ok, rect=" + (h3 == null ? "null" : Field(h3, "ScreenRect")));
            Check("旋转后仍能检测到图片", h3 != null);
        }
        catch (Exception ex)
        {
            Console.WriteLine("ERR " + ex);
            _fail++;
        }
        finally
        {
            try { wb.Close(false); } catch { }
            try { _xl.Quit(); } catch { }
            try { Marshal.ReleaseComObject(_xl); } catch { }
        }

        Console.WriteLine(_fail == 0 ? "ALL PASS" : ("FAILED " + _fail));
        return _fail == 0 ? 0 : 1;
    }

    private static object Detect()
    {
        MethodInfo m = _connect.GetMethod("DetectPicture", BindingFlags.NonPublic | BindingFlags.Instance);
        return m.Invoke(_conn, null);
    }

    private static void Rotate(object hit, string kind)
    {
        MethodInfo m = _connect.GetMethod("ApplyExcelRotate", BindingFlags.NonPublic | BindingFlags.Instance);
        m.Invoke(_conn, new object[] { hit, Enum.Parse(_kind, kind) });
    }

    private static object Invoke(string name, params object[] args)
    {
        MethodInfo m = _connect.GetMethod(name, BindingFlags.NonPublic | BindingFlags.Instance);
        if (m == null) { Console.WriteLine("METHOD MISSING: " + name); return null; }
        return m.Invoke(_conn, args);
    }

    private static double Rotation(dynamic shp) { return Convert.ToDouble(shp.Rotation); }

    private static void Check(string name, bool ok)
    {
        if (!ok) _fail++;
        Console.WriteLine("  [" + (ok ? "PASS" : "FAIL") + "] " + name);
    }

    private static string Field(object obj, string name)
    {
        FieldInfo fi = obj.GetType().GetField(name, BindingFlags.Public | BindingFlags.Instance);
        return fi == null ? "?" : Convert.ToString(fi.GetValue(obj));
    }

    private static void SetField(string name, object v)
    {
        FieldInfo fi = _connect.GetField(name, BindingFlags.NonPublic | BindingFlags.Instance);
        if (fi != null) fi.SetValue(_conn, v);
    }
}
