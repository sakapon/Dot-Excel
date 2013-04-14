using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;

namespace Keiho.Apps.DotExcel
{
    public static class ImageFile
    {
        public static Size GetSize(string filePath)
        {
            if (filePath == null) throw new ArgumentNullException("filePath");

            var bitmap = new Bitmap(filePath);

            return bitmap.Size;
        }

        public static IEnumerable<PointColor> Load(string filePath)
        {
            if (filePath == null) throw new ArgumentNullException("filePath");

            var bitmap = new Bitmap(filePath);

            return Enumerable.Range(0, bitmap.Height)
                .SelectMany(j => Enumerable.Range(0, bitmap.Width), (j, i) => new Point(i, j))
                .Select(p => new PointColor { Point = p, Color = bitmap.GetPixel(p.X, p.Y) })
                .Where(pc => !pc.Color.IsTransparent());
        }

        public static void SaveAsImage(this IEnumerable<PointColor> pointColors, string filePath, Size size)
        {
            if (pointColors == null) throw new ArgumentNullException("pointColors");
            if (filePath == null) throw new ArgumentNullException("filePath");

            var bitmap = new Bitmap(size.Width, size.Height);

            foreach (var item in pointColors)
            {
                bitmap.SetPixel(item.Point.X, item.Point.Y, item.Color);
            }
            bitmap.Save(filePath, GetImageFormat(filePath));
        }

        static ImageFormat GetImageFormat(string filePath)
        {
            switch (Path.GetExtension(filePath).ToLowerInvariant())
            {
                case ".png":
                    return ImageFormat.Png;
                case ".jpg":
                case ".jpeg":
                    return ImageFormat.Jpeg;
                case ".bmp":
                    return ImageFormat.Bmp;
                case ".gif":
                    return ImageFormat.Gif;
                case ".tif":
                case ".tiff":
                    return ImageFormat.Tiff;
                case ".ico":
                    return ImageFormat.Icon;
                default:
                    return ImageFormat.Png;
            }
        }
    }

    [DebuggerDisplay(@"\{({Point.X}, {Point.Y}), {Color.Name}\}")]
    public struct PointColor
    {
        public Point Point { get; set; }
        public Color Color { get; set; }
    }
}
