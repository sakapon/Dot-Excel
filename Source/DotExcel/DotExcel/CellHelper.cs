using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Keiho.Apps.DotExcel
{
    public static class CellHelper
    {
        const string Alphabets = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        static readonly Dictionary<char, int> AlphabetIndexes = Enumerable.Range(0, Alphabets.Length).ToDictionary(i => Alphabets[i], i => i);

        public static string ToColumnIndex(int index)
        {
            var section = ColumnSection.GetColumnSections().First(s => s.Contains(index));
            var sectionIndex = section.SectionIndex;
            var indexInSection = index - section.Start;

            var chars = Enumerable.Range(0, sectionIndex + 1).Reverse()
                .Select(i =>
                {
                    var p0 = (int)Math.Pow(Alphabets.Length, i);
                    var p1 = Alphabets.Length * p0;
                    return (indexInSection % p1) / p0;
                })
                .Select(i => Alphabets[i])
                .ToArray();

            return new string(chars);
        }

        public static int FromColumnIndex(string columnIndex)
        {
            if (columnIndex == null) throw new ArgumentNullException("columnIndex");
            if (!Regex.IsMatch(columnIndex, "^[A-Z]+$")) throw new ArgumentException("列名の形式に一致しません。", "columnIndex");

            var sectionIndex = columnIndex.Length - 1;
            var section = ColumnSection.GetColumnSections().First(s => s.SectionIndex == sectionIndex);

            var indexInSection = columnIndex
                .Select(c => AlphabetIndexes[c])
                .Select((q, i) => q * (int)Math.Pow(Alphabets.Length, sectionIndex - i))
                .Sum();

            return section.Start + indexInSection;
        }

        public static readonly Func<int, uint> ToRowIndex = i => (uint)(i + 1);
        public static readonly Func<uint, int> FromRowIndex = i => (int)i - 1;

        public static string ToCellReference(this Point point)
        {
            return ToColumnIndex(point.X) + Convert.ToString(ToRowIndex(point.Y));
        }

        public static Point ToPoint(this string cellReference)
        {
            if (cellReference == null) throw new ArgumentNullException("cellReference");

            var m = Regex.Match(cellReference, "^([A-Z]+)([1-9][0-9]*)$");
            if (!m.Success) throw new ArgumentException("セル名の形式に一致しません。", "cellReference");

            return new Point(FromColumnIndex(m.Groups[1].Value), FromRowIndex(Convert.ToUInt32(m.Groups[2].Value)));
        }

        public static string ToRowSpans(this Size size)
        {
            return "1:" + size.Width;
        }

        public static int ToWidthFromRowSpans(string rowSpansText)
        {
            if (rowSpansText == null) throw new ArgumentNullException("rowSpansText");

            var m = Regex.Match(rowSpansText, "^[0-9]+:([0-9]+)$");
            if (!m.Success) throw new ArgumentException("Spans.InnerText の形式に一致しません。", "rowSpansText");

            return Convert.ToInt32(m.Groups[1].Value);
        }
    }

    struct ColumnSection
    {
        public int SectionIndex { get; set; }
        public int Start { get; set; }
        public int Count { get; set; }

        public bool Contains(int value)
        {
            return Start <= value && value < Start + Count;
        }

        public static IEnumerable<ColumnSection> GetColumnSections()
        {
            var sectionSize = 1;
            var sum = 0;

            for (int i = 0; ; i++)
            {
                sectionSize *= 26;
                yield return new ColumnSection { SectionIndex = i, Start = sum, Count = sectionSize };
                sum += sectionSize;
            }
        }
    }
}
