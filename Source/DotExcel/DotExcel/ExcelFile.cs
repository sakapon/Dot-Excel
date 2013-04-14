using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Keiho.Apps.DotExcel.Consoles;
using TheColor = System.Drawing.Color;

namespace Keiho.Apps.DotExcel
{
    public static class ExcelFile
    {
        static readonly string SquareTemplateFilePath = Path.Combine(Path.GetDirectoryName(CommandUtility.GetEntryAssemblyPath()), "SquareTemplate.xlsx");

        public static Size GetSize(string filePath)
        {
            if (filePath == null) throw new ArgumentNullException("filePath");

            using (var document = SpreadsheetDocument.Open(filePath, false))
            {
                var worksheet = document.WorkbookPart.WorksheetParts.First().Worksheet;
                var sheetData = worksheet.Elements<SheetData>().First();

                var lastRow = sheetData.Elements<Row>().LastOrDefault();
                if (lastRow == null) return Size.Empty;

                return new Size(CellHelper.ToWidthFromRowSpans(lastRow.Spans.InnerText), (int)lastRow.RowIndex.Value);
            }
        }

        public static IEnumerable<PointColor> Load(string filePath)
        {
            if (filePath == null) throw new ArgumentNullException("filePath");

            using (var document = SpreadsheetDocument.Open(filePath, false))
            {
                var stylesheet = new ColorStylesheet(document.WorkbookPart.WorkbookStylesPart.Stylesheet);
                var worksheet = document.WorkbookPart.WorksheetParts.First().Worksheet;
                var sheetData = worksheet.Elements<SheetData>().First();

                return sheetData
                    .Descendants<Cell>()
                    .Select(c => new PointColor { Point = c.CellReference.Value.ToPoint(), Color = stylesheet.GetBackgroundColor(c.StyleIndex ?? 0) });
            }
        }

        public static void SaveAsExcel(this IEnumerable<PointColor> pointColors, string filePath, Size size)
        {
            if (pointColors == null) throw new ArgumentNullException("pointColors");
            if (filePath == null) throw new ArgumentNullException("filePath");

            throw new NotImplementedException().ToWarning();
        }
    }

    class ColorStylesheet
    {
        Stylesheet stylesheet;
        Dictionary<TheColor, uint> cellFormatIdsCache;
        Dictionary<uint, TheColor> colorsCache;

        public ColorStylesheet(Stylesheet stylesheet)
        {
            this.stylesheet = stylesheet;
            cellFormatIdsCache = new Dictionary<TheColor, uint>();
        }

        public UInt32Value AppendBackgroundColor(TheColor color)
        {
            if (cellFormatIdsCache.ContainsKey(color))
            {
                return cellFormatIdsCache[color];
            }

            var fill = new Fill(
                new PatternFill(
                    new ForegroundColor { Rgb = color.ToArgbString().ToUpperInvariant() },
                    new BackgroundColor { Indexed = 64 }
                ) { PatternType = PatternValues.Solid }
            );
            stylesheet.Fills.Append(fill);
            var fillId = stylesheet.Fills.Count++;

            var cellFormat = new CellFormat { NumberFormatId = 0, FontId = 0, FillId = fillId, BorderId = 0, FormatId = 0, ApplyFill = true };
            stylesheet.CellFormats.Append(cellFormat);
            var cellFormatId = stylesheet.CellFormats.Count++;

            cellFormatIdsCache[color] = cellFormatId;
            return cellFormatId;
        }

        public TheColor GetBackgroundColor(UInt32Value cellFormatId)
        {
            if (colorsCache == null)
            {
                var fillIdColorsCache = stylesheet.Fills.Elements<Fill>()
                    .Select((f, i) => new { FillId = i, Color = GetColor(f) })
                    .ToDictionary(x => (uint)x.FillId, x => x.Color);

                colorsCache = stylesheet.CellFormats.Elements<CellFormat>()
                    .Select((cf, i) => new { CellFormatId = i, cf.FillId })
                    .ToDictionary(x => (uint)x.CellFormatId, x => fillIdColorsCache[x.FillId]);
            }

            return colorsCache[cellFormatId];
        }

        static TheColor GetColor(Fill fill)
        {
            var patternFill = fill.PatternFill;
            return patternFill.PatternType != PatternValues.Solid ? TheColor.Transparent :
                patternFill.ForegroundColor == null ? TheColor.Transparent :
                patternFill.ForegroundColor.Rgb == null ? TheColor.Transparent :
                patternFill.ForegroundColor.Rgb.Value.ToColor();
        }
    }
}
