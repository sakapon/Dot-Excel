using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace Keiho.Apps.DotExcel
{
    public static class ColorHelper
    {
        public static Color ToColor(this string argb)
        {
            if (argb == null) throw new ArgumentNullException("argb");

            var argbInt = Convert.ToInt32(argb, 16);
            return Color.FromArgb(argbInt);
        }

        public static string ToArgbString(this Color color)
        {
            return Convert.ToString(color.ToArgb(), 16);
        }

        public static bool IsTransparent(this Color color)
        {
            return color.A == 0;
        }
    }
}
