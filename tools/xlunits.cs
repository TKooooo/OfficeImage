// 专门搞清楚两件事：
//   A) Excel 的 CropLeft 单位是"原图 100% 磅值"还是"当前显示磅值"（图片被 API 缩放过之后）
//   B) PictureResetAndSize 对被 API 缩放过 / 只裁过的图片分别还原到什么尺寸
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;

internal static class XlUnits
{
    private static void P(string k, object v) { Console.WriteLine(k + " = " + v); }
    private static string F(dynamic v) { return Convert.ToDouble(v).ToString("0.00"); }

    /// <summary>量一次"隐含原生宽度"：裁 10% 看宽度掉多少，再还原。</summary>
    private static string Probe(dynamic p)
    {
        try
        {
            dynamic pf = p.PictureFormat;
            double w0 = Convert.ToDouble(p.Width);
            double probe = w0 * 0.1;
            pf.CropLeft = probe;
            double w1 = Convert.ToDouble(p.Width);
            pf.CropLeft = 0;
            double native = (w1 > 0.01 && w1 < w0 * 0.999) ? probe / (1.0 - w1 / w0) : w0;
            return "cropDelta=" + (w0 - w1).ToString("0.00") + " impliedNative=" + native.ToString("0.0");
        }
        catch (Exception ex) { return "probe ERR " + ex.Message; }
    }

    [STAThread]
    private static int Main()
    {
        string png = Path.Combine(Path.GetTempPath(), "izunits.png");
        using (Bitmap bmp = new Bitmap(800, 600))
        using (Graphics g = Graphics.FromImage(bmp))
        {
            g.Clear(Color.Goldenrod);
            bmp.Save(png, ImageFormat.Png);
        }

        Type t = Type.GetTypeFromProgID("Excel.Application");
        dynamic xl = Activator.CreateInstance(t);
        dynamic wb = null;
        try
        {
            xl.Visible = false;
            xl.DisplayAlerts = false;
            wb = xl.Workbooks.Add();
            dynamic ws = wb.Worksheets.Item(1);

            // A) 原样插入（400x300）
            dynamic a = ws.Shapes.AddPicture(png, 0, -1, 100, 100, 400, 300);
            a.Name = "a";

            // B) 插入后 API 缩到 220x160
            dynamic b = ws.Shapes.AddPicture(png, 0, -1, 100, 500, 400, 300);
            b.Name = "b";
            b.Width = 220;
            b.Height = 160;

            P("a 400x300 ->", F(a.Width) + " x " + F(a.Height) + " (native 800x600px)");
            P("b resized  ->", F(b.Width) + " x " + F(b.Height));

            foreach (string nm in new string[] { "a", "b" })
            {
                dynamic p = ws.Shapes.Item(nm);
                dynamic pf = p.PictureFormat;
                double w0 = Convert.ToDouble(p.Width);
                pf.CropLeft = 22;
                double w1 = Convert.ToDouble(p.Width);
                pf.CropLeft = 0;
                P("crop22(" + nm + ")", F(w0) + " -> " + F(w1) + "  delta=" + (w0 - w1).ToString("0.00") +
                                        "  impliedNative=" + (22.0 / (1.0 - w1 / w0)).ToString("0.0"));
            }

            // B) 重设：先裁一刀再重设
            dynamic bb = ws.Shapes.Item("b");
            bb.PictureFormat.CropLeft = 20;
            bb.PictureFormat.CropTop = 10;
            P("b cropped  ->", F(bb.Width) + " x " + F(bb.Height));
            bb.Select();
            xl.CommandBars.ExecuteMso("PictureResetAndSize");
            P("b after reset", F(bb.Width) + " x " + F(bb.Height) + " crop=" +
                               bb.PictureFormat.CropLeft + "/" + bb.PictureFormat.CropTop);

            // a：只裁过没缩过
            dynamic aa = ws.Shapes.Item("a");
            aa.PictureFormat.CropRight = 40;
            P("a cropped  ->", F(aa.Width) + " x " + F(aa.Height));
            aa.Select();
            xl.CommandBars.ExecuteMso("PictureResetAndSize");
            P("a after reset", F(aa.Width) + " x " + F(aa.Height) + " crop=" +
                               aa.PictureFormat.CropRight + "/" + aa.PictureFormat.CropLeft);

            // C) 谁把"原生尺寸"改掉的？分别试 隐藏/显示、选中、SaveCopyAs
            dynamic c = ws.Shapes.AddPicture(png, 0, -1, 100, 900, 400, 300);
            c.Name = "c";
            c.Width = 220; c.Height = 160;
            c.Visible = 0;
            c.Visible = -1;
            dynamic cpf = c.PictureFormat;
            double cw0 = Convert.ToDouble(c.Width);
            cpf.CropLeft = 22;
            double cw1 = Convert.ToDouble(c.Width);
            cpf.CropLeft = 0;
            P("c hide/show then crop22", F(cw0) + " -> " + F(cw1) + " delta=" + (cw0 - cw1).ToString("0.00") +
                                          " impliedNative=" + (22.0 / (1.0 - cw1 / cw0)).ToString("0.0"));
            c.Select();
            xl.CommandBars.ExecuteMso("PictureResetAndSize");
            P("c after reset", F(c.Width) + " x " + F(c.Height));

            dynamic d = ws.Shapes.AddPicture(png, 0, -1, 500, 900, 400, 300);
            d.Name = "d";
            d.Width = 220; d.Height = 160;
            d.Select();
            dynamic dpf = d.PictureFormat;
            double dw0 = Convert.ToDouble(d.Width);
            dpf.CropLeft = 22;
            double dw1 = Convert.ToDouble(d.Width);
            dpf.CropLeft = 0;
            P("d select only then crop22", F(dw0) + " -> " + F(dw1) + " delta=" + (dw0 - dw1).ToString("0.00") +
                                            " impliedNative=" + (22.0 / (1.0 - dw1 / dw0)).ToString("0.0"));
            d.Select();
            xl.CommandBars.ExecuteMso("PictureResetAndSize");
            P("d after reset", F(d.Width) + " x " + F(d.Height));

            // E) 完整复刻加载项的流程，每一步量一次"隐含原生宽度"
            dynamic e = ws.Shapes.AddPicture(png, 0, -1, 900, 900, 400, 300);
            e.Name = "e";
            e.Width = 220; e.Height = 160;
            P("e[resize]", F(e.Width) + " " + Probe(e));
            string cp = Path.Combine(Path.GetTempPath(), "izunits_copy.xlsx");
            if (File.Exists(cp)) File.Delete(cp);
            wb.SaveCopyAs(cp);
            P("e[savecopy]", Probe(e));
            dynamic es = ws.Shapes.AddShape(92, e.Left, e.Top, e.Width, e.Height);
            es.Fill.UserPicture(png);
            es.Name = "e#shape";
            e.Visible = 0;
            P("e[shape+hide]", Probe(e) + " visible=" + e.Visible);
            es.Delete();
            e.Visible = -1;
            e.Select();
            P("e[restore]", F(e.Width) + " " + Probe(e));
            e.Select();
            xl.CommandBars.ExecuteMso("PictureResetAndSize");
            P("e[reset]", F(e.Width) + " x " + F(e.Height));
            if (File.Exists(cp)) File.Delete(cp);
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
}
