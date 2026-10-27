using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Reflection;
using System.Windows.Forms;

/// <summary>
/// 把三个浮动按钮按 6 倍像素放大渲染成对照图（第一行压在深蓝底上，模拟图片背景；
/// 第二行白底常态；第三行白底悬停），用来目检圆角、描边、图标线条的精细程度。
/// 直接取 MagnifierButton.Compose() 的真实输出 —— 分层窗口的内容就是它，没有第二条绘制路径。
/// </summary>
class BtnProbe
{
    private const int Scale = 6;

    [STAThread]
    static void Main(string[] args)
    {
        string dll = args[0];
        string outPng = args.Length > 1 ? args[1] : "buttons.png";
        Assembly asm = Assembly.LoadFrom(dll);

        Type btnType = asm.GetType("ImageZoomAddin.MagnifierButton");
        Type kindType = asm.GetType("ImageZoomAddin.ButtonKind");
        string[] kinds = { "Crop", "Zoom", "Rotate" };

        MethodInfo compose = btnType.GetMethod("Compose", BindingFlags.NonPublic | BindingFlags.Instance);
        FieldInfo hoverField = btnType.GetField("_hover", BindingFlags.NonPublic | BindingFlags.Instance);
        if (compose == null || hoverField == null) { Console.WriteLine("MISSING Compose/_hover"); return; }

        int cell = 30 * Scale;
        int pad = 14;
        using (Bitmap sheet = new Bitmap(3 * (cell + pad) + pad, 3 * (cell + pad) + pad, PixelFormat.Format32bppArgb))
        using (Graphics g = Graphics.FromImage(sheet))
        {
            g.Clear(Color.White);
            using (SolidBrush b = new SolidBrush(Color.FromArgb(38, 106, 205)))
                g.FillRectangle(b, 0, 0, sheet.Width, cell + 2 * pad);

            g.InterpolationMode = InterpolationMode.NearestNeighbor;
            g.PixelOffsetMode = PixelOffsetMode.Half;

            for (int row = 0; row < 3; row++)
            {
                bool hover = row == 2;          // 第一行蓝底常态，第二行白底常态，第三行白底悬停
                for (int i = 0; i < kinds.Length; i++)
                {
                    object btn = Activator.CreateInstance(btnType,
                        new object[] { Enum.Parse(kindType, kinds[i]) });
                    hoverField.SetValue(btn, hover);

                    Bitmap shot = (Bitmap)compose.Invoke(btn, null);
                    int x = pad + i * (cell + pad);
                    int y = pad + row * (cell + pad);
                    g.DrawImage(shot, new Rectangle(x, y, cell, cell));
                    shot.Dispose();
                    ((Form)btn).Dispose();
                }
            }

            sheet.Save(outPng, ImageFormat.Png);
        }
        Console.WriteLine("saved " + outPng);
    }
}
