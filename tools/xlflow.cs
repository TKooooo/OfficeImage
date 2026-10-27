// Excel 裁剪流程干跑：不依赖加载项 GUI，验证 ApplyExcelShapeCrop / Restore / Reset 依赖的每一个 COM 行为
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Xml.Linq;

internal static class XlFlow
{
    private static void P(string k, object v) { Console.WriteLine(k + " = " + v); }

    [STAThread]
    private static int Main()
    {
        string png = Path.Combine(Path.GetTempPath(), "izflow_src.png");
        using (Bitmap bmp = new Bitmap(800, 600))
        using (Graphics g = Graphics.FromImage(bmp))
        {
            g.Clear(Color.SeaGreen);
            g.DrawRectangle(Pens.White, 10, 10, 780, 580);
            bmp.Save(png, ImageFormat.Png);
        }

        Type t = Type.GetTypeFromProgID("Excel.Application");
        dynamic xl = Activator.CreateInstance(t);
        dynamic wb = null, ws = null;
        try
        {
            xl.Visible = false;
            xl.DisplayAlerts = false;
            wb = xl.Workbooks.Add();
            ws = wb.Worksheets.Item(1);

            // ---------- 1) 原图
            dynamic pic = ws.Shapes.AddPicture(png, 0, -1, 100, 100, 400, 300);
            pic.Name = "img1";
            P("pic", pic.Name + " " + pic.Left + "," + pic.Top + " " + pic.Width + "x" + pic.Height);

            // ---------- 2) 裁剪为形状（模拟 ApplyExcelShapeCrop）
            dynamic sp = ws.Shapes.AddShape(92, pic.Left, pic.Top, pic.Width, pic.Height);   // 92 = 5pointStar
            sp.Fill.UserPicture(png);
            sp.Line.Visible = 0;
            P("sp.Fill.Type", sp.Fill.Type);
            sp.AlternativeText = "IZADDIN:SHAPE:star5|img1";
            sp.Name = "img1#shape";
            pic.Visible = 0;
            sp.Select();

            dynamic sel = xl.Selection;
            P("selected name", sel.ShapeRange.Item(1).Name);
            P("selected alt", sel.ShapeRange.Item(1).AlternativeText);
            P("selected type", sel.ShapeRange.Item(1).Type);
            P("pic visible after hide", pic.Visible);
            P("sp placement", sp.Placement);

            // ---------- 3) 找回原图（模拟 FindOriginalPicture）
            string orig = Convert.ToString(sp.AlternativeText);
            orig = orig.Substring(orig.LastIndexOf('|') + 1);
            dynamic found = null;
            try { found = ws.Shapes.Item(orig); } catch (Exception ex) { P("Shapes.Item", "ERR " + ex.Message); }
            P("found original", found == null ? "null" : Convert.ToString(found.Name));

            // ---------- 4) 落盘后按名字能否挖出图片字节（模拟 Ooxml.FromExcelPackage + sp 分支）
            string tmp = Path.Combine(Path.GetTempPath(), "izflow_out.xlsx");
            if (File.Exists(tmp)) File.Delete(tmp);
            wb.SaveCopyAs(tmp);
            ProbePackage(tmp, "img1#shape");
            ProbePackage(tmp, "img1");

            // ---------- 5) 还原成图片（模拟 RestoreExcelPicture）
            sp.Delete();
            found.Visible = -1;
            found.Select();
            P("after restore visible", found.Visible);
            P("after restore selected", xl.Selection.ShapeRange.Item(1).Name);

            // ---------- 6) 按比例裁剪（16:9）后能否用 PictureResetAndSize 还原
            dynamic pf = found.PictureFormat;
            double Wf = Convert.ToDouble(found.Width), Hf = Convert.ToDouble(found.Height);
            double probe = Wf * 0.1;
            pf.CropLeft = probe;
            double Wp = Convert.ToDouble(found.Width);
            pf.CropLeft = 0;
            double W0 = (Wp > 0.01 && Wp < Wf * 0.999) ? probe / (1.0 - Wp / Wf) : Wf;
            P("native width W0", W0.ToString("0.0") + " (current " + Wf + "x" + Hf + ")");

            double R = 16.0 / 9.0;
            for (int k = 0; k < 3; k++)
            {
                double Wv = Convert.ToDouble(found.Width), Hv = Convert.ToDouble(found.Height);
                double cur = Wv / Hv;
                if (Math.Abs(cur - R) <= R * 0.0015) break;
                if (cur > R) { double dc = (Wv - Hv * R) * W0 / Wf / 2.0; pf.CropLeft = Convert.ToDouble(pf.CropLeft) + dc; pf.CropRight = Convert.ToDouble(pf.CropRight) + dc; }
                else { double dc = (Hv - Wv / R) * (W0 * Hf / Wf) / Hf / 2.0; pf.CropTop = Convert.ToDouble(pf.CropTop) + dc; pf.CropBottom = Convert.ToDouble(pf.CropBottom) + dc; }
            }
            P("after 16:9", Convert.ToDouble(found.Width).ToString("0.0") + "x" + Convert.ToDouble(found.Height).ToString("0.0") +
                            " ratio=" + (Convert.ToDouble(found.Width) / Convert.ToDouble(found.Height)).ToString("0.0000"));

            // ---------- 7) 原生重设
            try
            {
                found.Select();
                xl.CommandBars.ExecuteMso("PictureResetAndSize");
                P("after PictureResetAndSize", Convert.ToDouble(found.Width).ToString("0.0") + "x" +
                                               Convert.ToDouble(found.Height).ToString("0.0") +
                                               " crop=" + found.PictureFormat.CropLeft + "/" + found.PictureFormat.CropTop);
            }
            catch (Exception ex) { P("PictureResetAndSize", "ERR " + ex.Message); }

            // ---------- 8) 原生裁剪命令（菜单里的「自由裁剪」）
            try
            {
                found.Select();
                P("pressed before", xl.CommandBars.GetPressedMso("PictureCrop"));
                xl.CommandBars.ExecuteMso("PictureCrop");
                P("pressed after", xl.CommandBars.GetPressedMso("PictureCrop"));
                P("PictureCrop", "ok");
            }
            catch (Exception ex) { P("PictureCrop", "ERR " + ex.Message); }

            if (File.Exists(tmp)) File.Delete(tmp);
        }
        catch (Exception ex) { Console.WriteLine("FATAL " + ex); }
        finally
        {
            try { if (wb != null) wb.Close(false); } catch { }
            try { xl.Quit(); } catch { }
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        Console.WriteLine("DONE");
        return 0;
    }

    private static void ProbePackage(string xlsx, string shapeName)
    {
        try
        {
            using (ZipArchive zip = ZipFile.OpenRead(xlsx))
            {
                foreach (ZipArchiveEntry e in zip.Entries)
                {
                    if (!e.FullName.StartsWith("xl/drawings/drawing") || !e.FullName.EndsWith(".xml")) continue;
                    XDocument d;
                    using (Stream s = e.Open()) d = XDocument.Load(s);
                    foreach (XElement cnv in d.Descendants().Where(x => x.Name.LocalName == "cNvPr"))
                    {
                        if (!string.Equals((string)cnv.Attribute("name"), shapeName, StringComparison.OrdinalIgnoreCase)) continue;
                        XElement scope = cnv.Ancestors().FirstOrDefault(a => a.Name.LocalName == "pic");
                        if (scope == null) scope = cnv.Ancestors().FirstOrDefault(a => a.Name.LocalName == "sp");
                        if (scope == null) { P("  scope(" + shapeName + ")", "none"); continue; }
                        XElement blip = scope.Descendants().FirstOrDefault(x => x.Name.LocalName == "blip");
                        XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                        string rid = blip == null ? null : (string)blip.Attribute(r + "embed");
                        P("  embed(" + shapeName + ")", (scope.Name.LocalName) + " -> " + (rid == null ? "no blip" : rid));
                        if (rid == null) continue;
                        string relsPath = "xl/drawings/_rels/" + Path.GetFileName(e.FullName) + ".rels";
                        ZipArchiveEntry rels = zip.GetEntry(relsPath);
                        if (rels == null) { P("  rels", "missing"); continue; }
                        XDocument rd;
                        using (Stream s = rels.Open()) rd = XDocument.Load(s);
                        XElement rel = rd.Descendants().FirstOrDefault(x => x.Name.LocalName == "Relationship" && (string)x.Attribute("Id") == rid);
                        P("  target", rel == null ? "?" : (string)rel.Attribute("Target"));
                    }
                }
            }
        }
        catch (Exception ex) { P("ProbePackage", "ERR " + ex.Message); }
    }
}
