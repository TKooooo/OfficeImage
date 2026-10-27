using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;

/// <summary>
/// A/B 对照图：第一行是 WPS 原图里的三个按钮（各 30x30 裁出来放大 6 倍），
/// 第二行是本插件渲染的三个按钮。用来判断"像不像"。
/// </summary>
class Compare
{
    static void Main(string[] a)
    {
        string wpsPng = a[0];      // WPS 截图
        string minePng = a[1];     // btnprobe 渲染的对照图
        string outPng = a[2];

        int scale = 6, cell = 30 * scale, pad = 12;
        int[] wpsX = { 477, 477, 477 };
        int[] wpsY = { 167, 202, 237 };

        using (Bitmap sheet = new Bitmap(3 * (cell + pad) + pad, 2 * (cell + pad) + pad, PixelFormat.Format32bppArgb))
        using (Graphics g = Graphics.FromImage(sheet))
        {
            g.Clear(Color.FromArgb(238, 240, 244));
            g.InterpolationMode = InterpolationMode.NearestNeighbor;
            g.PixelOffsetMode = PixelOffsetMode.Half;

            using (Image wps = Image.FromFile(wpsPng))
            using (Image mine = Image.FromFile(minePng))
            {
                for (int i = 0; i < 3; i++)
                {
                    g.DrawImage(wps,
                        new Rectangle(pad + i * (cell + pad), pad, cell, cell),
                        new Rectangle(wpsX[i], wpsY[i], 30, 30), GraphicsUnit.Pixel);

                    Rectangle srcMine = new Rectangle(14 + i * (180 + 14), 14, 180, 180);
                    g.DrawImage(mine,
                        new Rectangle(pad + i * (cell + pad), pad + cell + pad, cell, cell),
                        srcMine, GraphicsUnit.Pixel);
                }
            }

            sheet.Save(outPng, ImageFormat.Png);
        }
        Console.WriteLine("saved " + outPng);
    }
}
