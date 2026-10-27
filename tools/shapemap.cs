// 一次性探测：Excel 的 msoAutoShapeType 数值 -> drawing xml 里的 a:prstGeom prst 名字
// 做法：在隐藏的 Excel 实例里为 1..250 每个数值建一个形状，落盘后解析 drawing1.xml 反查。
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

internal static class ShapeMap
{
    [STAThread]
    private static int Main()
    {
        Type t = Type.GetTypeFromProgID("Excel.Application");
        dynamic xl = Activator.CreateInstance(t);
        dynamic wb = null, ws = null;
        try
        {
            xl.Visible = false;
            xl.DisplayAlerts = false;
            wb = xl.Workbooks.Add();
            ws = wb.Worksheets.Item(1);

            int made = 0;
            for (int n = 1; n <= 250; n++)
            {
                try
                {
                    int left = 10 + (n % 20) * 20;
                    int top = 10 + (n / 20) * 20;
                    dynamic shp = ws.Shapes.AddShape(n, left, top, 18, 14);
                    shp.Name = "t" + n;
                    made++;
                }
                catch { }
            }
            Console.WriteLine("shapes created = " + made);

            string tmp = Path.Combine(Path.GetTempPath(), "izshapemap.xlsx");
            if (File.Exists(tmp)) File.Delete(tmp);
            wb.SaveCopyAs(tmp);

            Dictionary<string, string> map = new Dictionary<string, string>();
            using (ZipArchive zip = ZipFile.OpenRead(tmp))
            {
                XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
                foreach (ZipArchiveEntry e in zip.Entries)
                {
                    if (!e.FullName.StartsWith("xl/drawings/drawing") || !e.FullName.EndsWith(".xml")) continue;
                    XDocument d;
                    using (Stream s = e.Open()) d = XDocument.Load(s);
                    foreach (XElement el in d.Descendants())
                    {
                        if (el.Name.LocalName != "cNvPr") continue;
                        string name = (string)el.Attribute("name");
                        if (name == null || !name.StartsWith("t")) continue;
                        XElement sp = el.Ancestors().FirstOrDefault(x => x.Name.LocalName == "sp");
                        if (sp == null) continue;
                        XElement geom = sp.Descendants().FirstOrDefault(x => x.Name.LocalName == "prstGeom");
                        string prst = geom == null ? "?" : (string)geom.Attribute("prst");
                        if (!map.ContainsKey(name)) map[name] = prst;
                    }
                }
            }

            // 输出：数值  prst（按 prst 排序便于人工核对）
            List<KeyValuePair<string, string>> rows = new List<KeyValuePair<string, string>>();
            foreach (KeyValuePair<string, string> kv in map)
            {
                int n;
                if (!int.TryParse(kv.Key.Substring(1), out n)) continue;
                rows.Add(new KeyValuePair<string, string>(kv.Value, n.ToString()));
            }
            rows.Sort(delegate (KeyValuePair<string, string> x, KeyValuePair<string, string> y)
            {
                int c = string.CompareOrdinal(x.Key, y.Key);
                if (c != 0) return c;
                return int.Parse(x.Value).CompareTo(int.Parse(y.Value));
            });
            foreach (KeyValuePair<string, string> kv in rows)
                Console.WriteLine(kv.Key.PadRight(30) + " = " + kv.Value);

            File.Delete(tmp);
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
