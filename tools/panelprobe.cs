using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Reflection;
using System.Windows.Forms;

internal static class PanelProbe
{
    private static Type _type;

    [STAThread]
    private static void Main(string[] args)
    {
        string dll = args.Length > 0 ? args[0] : "ImageZoomAddin.dll";
        Assembly asm = Assembly.LoadFrom(dll);
        _type = asm.GetType("ImageZoomAddin.CropMenuForm");
        if (_type == null) { Console.WriteLine("TYPE NOT FOUND"); return; }

        object o = Activator.CreateInstance(_type, true);
        Form f = (Form)o;
        _type.GetMethod("Relayout").Invoke(o, null);

        f.Show();
        Pump();
        Save(f, "panel-shape.png");

        SetTab(o, 1);
        Pump();
        Save(f, "panel-ratio.png");

        // 模拟悬停：把鼠标移到某个形状格子上（只改 _hover，避免真要移动真鼠标）
        SetTab(o, 0);
        Pump();
        SetField(o, "_hover", 2);
        Invoke("Invalidate");
        Pump();
        Save(f, "panel-hover.png");

        f.Close();
        Console.WriteLine("done");
    }

    private static void SetTab(object o, int tab)
    {
        SetField(o, "_tab", tab);
        _type.GetMethod("Relayout").Invoke(o, null);
        Invoke("Invalidate");
    }

    private static void SetField(object o, string name, object v)
    {
        FieldInfo fi = _type.GetField(name, BindingFlags.NonPublic | BindingFlags.Instance);
        if (fi != null) fi.SetValue(o, v);
    }

    private static void Invoke(string name)
    {
    }

    private static void Pump()
    {
        for (int i = 0; i < 12; i++)
        {
            Application.DoEvents();
            System.Threading.Thread.Sleep(20);
        }
    }

    private static void Save(Form f, string path)
    {
        Application.DoEvents();
        Bitmap bmp = new Bitmap(f.Width, f.Height);
        f.DrawToBitmap(bmp, new Rectangle(0, 0, f.Width, f.Height));
        bmp.Save(path, ImageFormat.Png);
        Console.WriteLine(path + "  " + f.Width + "x" + f.Height);
        bmp.Dispose();
    }
}
