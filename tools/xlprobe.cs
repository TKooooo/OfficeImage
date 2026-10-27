// Excel 能力探针：验证本加载项要用的几个 Excel COM 行为
// 编译：见 tools/xlprobe-build.bat（输出 tools/_probe/xlprobe.exe）
// 运行：独立 Excel 实例，不碰用户已打开的窗口
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;

internal static class XlProbe
{
    private static void P(string k, object v) { Console.WriteLine(k + " = " + v); }

    [STAThread]
    private static int Main()
    {
        string png = Path.Combine(Path.GetTempPath(), "izprobe_src.png");
        MakePng(png);
        P("png", png + " (" + new FileInfo(png).Length + " bytes)");

        Type t = Type.GetTypeFromProgID("Excel.Application");
        if (t == null) { Console.WriteLine("Excel not installed"); return 2; }
        dynamic xl = Activator.CreateInstance(t);
        dynamic wb = null, ws = null, pic = null, shp = null;

        try
        {
            xl.Visible = false;
            xl.DisplayAlerts = false;
            wb = xl.Workbooks.Add();
            ws = wb.Worksheets.Item(1);

            // ---- 1) 图片裁剪对尺寸的影响
            pic = ws.Shapes.AddPicture(png, 0, -1, 100, 100, 400, 300);
            pic.Name = "probe_pic";
            P("pic Type/Left/Top/W/H", pic.Type + " / " + pic.Left + " / " + pic.Top + " / " + pic.Width + " / " + pic.Height);

            dynamic pf = pic.PictureFormat;
            P("pf.CropLeft(init)", pf.CropLeft);
            pf.CropLeft = 30;
            P("after CropLeft=30 -> W/H/Left", pic.Width + " / " + pic.Height + " / " + pic.Left);
            pf.CropLeft = 0;
            P("back to 0 -> W/H/Left", pic.Width + " / " + pic.Height + " / " + pic.Left);

            // ---- 2) 原生 MSO 命令在 Excel 里是否可用
            pic.Select();
            try { P("GetEnabledMso(PictureCrop)", xl.CommandBars.GetEnabledMso("PictureCrop")); }
            catch (Exception ex) { P("GetEnabledMso(PictureCrop)", "ERR " + ex.Message); }
            try { P("GetEnabledMso(PictureResetAndSize)", xl.CommandBars.GetEnabledMso("PictureResetAndSize")); }
            catch (Exception ex) { P("GetEnabledMso(PictureResetAndSize)", "ERR " + ex.Message); }

            // ---- 3) 自选图形 + 图片填充
            shp = ws.Shapes.AddShape(1, 120, 450, 400, 300);
            shp.Name = "probe_shape";
            P("shape Type", shp.Type);
            P("shape Fill.Type(before)", shp.Fill.Type);
            try { shp.Fill.UserPicture(png); P("Fill.UserPicture", "ok"); }
            catch (Exception ex) { P("Fill.UserPicture", "ERR " + ex.Message); }
            P("shape Fill.Type(after)", shp.Fill.Type);
            try { shp.Line.Visible = 0; P("Line.Visible=0", "ok"); } catch (Exception ex) { P("Line.Visible=0", "ERR " + ex.Message); }
            try { shp.AlternativeText = "IZADDIN:SHAPE:ellipse"; P("AlternativeText", shp.AlternativeText); }
            catch (Exception ex) { P("AlternativeText", "ERR " + ex.Message); }
            try { P("shape.PictureFormat", shp.PictureFormat == null ? "null" : "present"); }
            catch (Exception ex) { P("shape.PictureFormat", "ERR " + ex.Message); }

            // ---- 4) Visible 开关（用来"藏"原图）
            pic.Visible = 0;
            P("pic.Visible=0", pic.Visible);
            pic.Visible = -1;
            P("pic.Visible=-1", pic.Visible);

            // ---- 5) 形状改名 / 尺寸可写
            try { shp.Left = 130; P("shape.Left writable", shp.Left); } catch (Exception ex) { P("shape.Left writable", "ERR " + ex.Message); }

            // ---- 6) 落盘后看 drawing xml 里 sp 的 blipFill
            string tmp = Path.Combine(Path.GetTempPath(), "izprobe_out.xlsx");
            if (File.Exists(tmp)) File.Delete(tmp);
            wb.SaveCopyAs(tmp);
            P("saved", File.Exists(tmp) + " " + (File.Exists(tmp) ? new FileInfo(tmp).Length : 0) + " bytes");
            DumpDrawing(tmp);
        }
        catch (Exception ex)
        {
            Console.WriteLine("FATAL " + ex);
        }
        finally
        {
            try { if (wb != null) wb.Close(false); } catch (Exception ex) { P("wb.Close", "ERR " + ex.Message); }
            try { xl.Quit(); } catch (Exception ex) { P("xl.Quit", "ERR " + ex.Message); }
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        Console.WriteLine("DONE");
        return 0;
    }

    private static void MakePng(string path)
    {
        using (Bitmap bmp = new Bitmap(800, 600))
        using (Graphics g = Graphics.FromImage(bmp))
        {
            g.Clear(Color.Orange);
            g.DrawRectangle(Pens.Blue, 10, 10, 780, 580);
            bmp.Save(path, ImageFormat.Png);
        }
    }

    private static void DumpDrawing(string xlsx)
    {
        try
        {
            using (ZipArchive zip = ZipFile.OpenRead(xlsx))
            {
                foreach (ZipArchiveEntry e in zip.Entries)
                {
                    if (!e.FullName.StartsWith("xl/drawings/drawing") || !e.FullName.EndsWith(".xml")) continue;
                    using (Stream s = e.Open())
                    using (StreamReader r = new StreamReader(s))
                    {
                        string xml = r.ReadToEnd();
                        P("drawing", e.FullName + " len=" + xml.Length);
                        int i = xml.IndexOf("blipFill", StringComparison.Ordinal);
                        if (i < 0) { P("  blipFill", "not found"); continue; }
                        int from = Math.Max(0, i - 700);
                        Console.WriteLine("  ..." + xml.Substring(from, Math.Min(1100, xml.Length - from)) + "...");
                    }
                }
                foreach (ZipArchiveEntry e in zip.Entries)
                    if (e.FullName.StartsWith("xl/media/")) P("media", e.FullName + " " + e.Length);
            }
        }
        catch (Exception ex) { P("DumpDrawing", "ERR " + ex.Message); }
    }
}
