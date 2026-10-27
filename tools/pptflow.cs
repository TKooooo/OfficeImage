// PowerPoint 的三个功能到底怎么落地：
//   裁剪为形状 → Shape.AutoShapeType 可写吗？写完 PictureFormat 还能用吗？能设回矩形吗？
//   重设      → PictureReset / PictureResetAndSize 执行后形状、裁剪、尺寸各变成什么
//   识别      → 占位符里的图片 Type 是多少？多选时 Selection 长什么样？
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;

internal static class PptFlow
{
    private static void P(string k, object v) { Console.WriteLine("  " + k + " = " + v); }
    private static string F(dynamic v) { try { return Convert.ToDouble(v).ToString("0.00"); } catch { return "?"; } }

    private static string State(dynamic shp)
    {
        try
        {
            return "type=" + shp.Type + " ast=" + shp.AutoShapeType +
                   " L=" + F(shp.Left) + " T=" + F(shp.Top) + " W=" + F(shp.Width) + " H=" + F(shp.Height) +
                   " rot=" + F(shp.Rotation) +
                   " crop=" + F(shp.PictureFormat.CropLeft) + "/" + F(shp.PictureFormat.CropRight) + "/" +
                   F(shp.PictureFormat.CropTop) + "/" + F(shp.PictureFormat.CropBottom);
        }
        catch (Exception ex) { return "ERR " + ex.Message; }
    }

    [STAThread]
    private static int Main()
    {
        string png = Path.Combine(Path.GetTempPath(), "izflow.png");
        using (Bitmap bmp = new Bitmap(800, 450))
        using (Graphics g = Graphics.FromImage(bmp)) { g.Clear(Color.Magenta); bmp.Save(png, ImageFormat.Png); }

        Type t = Type.GetTypeFromProgID("PowerPoint.Application");
        dynamic app = Activator.CreateInstance(t);
        dynamic prs = null;
        try
        {
            app.Visible = true;
            prs = app.Presentations.Add();
            dynamic slide = prs.Slides.Add(1, 12);

            dynamic pic = slide.Shapes.AddPicture(png, 0, -1, 200, 200, 600, 337.5);
            pic.Name = "flowpic";
            pic.Select();
            Thread.Sleep(400);

            // ---------- 1. 裁剪为形状 ----------
            Console.WriteLine("==== 1. crop to shape ====");
            P("start", State(pic));
            string e0 = Path.Combine(Path.GetTempPath(), "izflow_rect.png");
            string e1 = Path.Combine(Path.GetTempPath(), "izflow_oval.png");
            string e2 = Path.Combine(Path.GetTempPath(), "izflow_star.png");
            try { pic.Export(e0, 2); } catch (Exception ex) { P("export0", ex.Message); }

            try
            {
                pic.AutoShapeType = 9;    // msoShapeOval
                Thread.Sleep(300);
                P("after ast=9", State(pic));
                try { pic.Export(e1, 2); } catch (Exception ex) { P("export1", ex.Message); }
                P("选中还在?", "selection.Type=" + app.ActiveWindow.Selection.Type +
                               " count=" + app.ActiveWindow.Selection.ShapeRange.Count);
            }
            catch (Exception ex) { P("ast=9 FAIL", ex.Message); }

            // 形状裁剪后再改裁剪值，还灵吗
            try
            {
                dynamic pf = pic.PictureFormat;
                pf.CropLeft = 30; Thread.Sleep(250);
                P("oval + CropLeft=30", State(pic));
                pf.CropLeft = 0; Thread.Sleep(250);
            }
            catch (Exception ex) { P("oval+crop FAIL", ex.Message); }

            try
            {
                pic.AutoShapeType = 92;   // msoShape5pointStar
                Thread.Sleep(300);
                P("after ast=92(star)", State(pic));
                try { pic.Export(e2, 2); } catch (Exception ex) { P("export2", ex.Message); }
            }
            catch (Exception ex) { P("ast=92 FAIL", ex.Message); }

            // 导出图的角像素：真的裁掉了？(透明/白)
            foreach (string f in new[] { e0, e1, e2 })
            {
                if (!File.Exists(f)) { P("file", f + " missing"); continue; }
                using (Bitmap b = new Bitmap(f))
                {
                    Color c = b.GetPixel(3, 3);
                    Color m = b.GetPixel(b.Width / 2, b.Height / 2);
                    P(Path.GetFileName(f), b.Width + "x" + b.Height + "  corner=" + c.Name + "(A=" + c.A + ")  center=" + m.Name);
                }
            }

            // ---------- 2. 重设 ----------
            Console.WriteLine("==== 2. reset ====");
            pic.Select(); Thread.Sleep(200);
            P("before reset", State(pic));
            foreach (string id in new[] { "PictureResetAndSize", "PictureReset" })
            {
                try
                {
                    P(id + " enabled", app.CommandBars.GetEnabledMso(id));
                    app.CommandBars.ExecuteMso(id);
                    Thread.Sleep(600);
                    P("after " + id, State(pic));
                }
                catch (Exception ex) { P(id + " FAIL", ex.Message); }
            }

            // ---------- 3. 占位符里的图片 ----------
            Console.WriteLine("==== 3. placeholder ====");
            try
            {
                dynamic s2 = prs.Slides.Add(2, 1);    // ppLayoutTitle
                Thread.Sleep(300);
                dynamic ph = null;
                int n = s2.Shapes.Count;
                P("layout1 shapes", n);
                for (int i = 1; i <= n; i++)
                {
                    dynamic sh = s2.Shapes.Item(i);
                    P("  shape" + i, "name=" + sh.Name + " type=" + sh.Type +
                                    " phType=" + SafeInt(delegate { return sh.PlaceholderFormat.Type; }));
                    if (SafeInt(delegate { return sh.PlaceholderFormat.Type; }) == 2) ph = sh;   // ppPlaceholderObject=7? 先打印
                }
                if (ph != null)
                {
                    P("content placeholder found", ph.Name);
                    try
                    {
                        ph.Select();
                        Thread.Sleep(200);
                        ph.Fill.UserPicture(png);      // 图片填充占位符
                        Thread.Sleep(300);
                        P("after UserPicture", "type=" + ph.Type + " fillType=" + SafeInt(delegate { return ph.Fill.Type; }));
                    }
                    catch (Exception ex) { P("UserPicture FAIL", ex.Message); }
                }
            }
            catch (Exception ex) { P("placeholder section FAIL", ex.Message); }

            // ---------- 4. 多选 / 组合 ----------
            Console.WriteLine("==== 4. multi-select ====");
            try
            {
                dynamic s3 = prs.Slides.Add(3, 12);
                Thread.Sleep(200);
                dynamic a = s3.Shapes.AddPicture(png, 0, -1, 100, 100, 200, 112.5);
                a.Name = "a1";
                dynamic b = s3.Shapes.AddPicture(png, 0, -1, 400, 100, 200, 112.5);
                b.Name = "b2";
                app.ActiveWindow.View.GotoSlide(3);
                Thread.Sleep(400);
                a.Select();
                b.Select(0);   // msoFalse -> 追加选择
                Thread.Sleep(300);
                P("multi Selection.Type", app.ActiveWindow.Selection.Type);
                P("multi ShapeRange.Count", app.ActiveWindow.Selection.ShapeRange.Count);
                P("multi ShapeRange[1]", app.ActiveWindow.Selection.ShapeRange[1].Name);

                // 组合
                dynamic grp = app.ActiveWindow.Selection.ShapeRange.Group();
                Thread.Sleep(400);
                P("group Type", grp.Type);
                P("group Selection.Type", app.ActiveWindow.Selection.Type);
                P("group ShapeRange.Count", app.ActiveWindow.Selection.ShapeRange.Count);
            }
            catch (Exception ex) { P("multi FAIL", ex.Message); }

            // ---------- 5. 原始字节：包内提取 ----------
            Console.WriteLine("==== 5. package ====");
            try
            {
                string copy = Path.Combine(Path.GetTempPath(), "izflow.pptx");
                prs.SaveCopyAs(copy);
                P("SaveCopyAs bytes", new FileInfo(copy).Length);
                File.Delete(copy);
            }
            catch (Exception ex) { P("SaveCopyAs FAIL", ex.Message); }

            // ---------- 6. 视图类型 ----------
            Console.WriteLine("==== 6. view ====");
            try { app.ActiveWindow.ViewType = 7; Thread.Sleep(400); P("slideSorter ViewType", app.ActiveWindow.ViewType); } catch (Exception ex) { P("sorter", ex.Message); }
            try { app.ActiveWindow.ViewType = 9; Thread.Sleep(400); P("normal ViewType", app.ActiveWindow.ViewType); } catch { }
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

    private static int SafeInt(Func<object> f)
    {
        try { return Convert.ToInt32(f()); } catch { return -999; }
    }
}
