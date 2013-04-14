using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Keiho.Apps.DotExcel.Consoles;

namespace Keiho.Apps.DotExcel
{
    [CommandsClass]
    [CommandUsage(@"
Usage:
dotexcel ""file1"" [""file2"" ...] [/format:png|jpg|jpeg|bmp|gif|tif|tiff|ico]
dotexcel toexcel ""file1"" [""file2"" ...]
dotexcel toimage ""file1"" [""file2"" ...] [/format:png|jpg|jpeg|bmp|gif|tif|tiff|ico]
dotexcel help
")]
    public static class Commands
    {
        internal static readonly Func<string[], string[]> CorrectArgs =
            args =>
                args.Length == 0 ? args :
                Regex.IsMatch(args[0], "^.*\\.(png|jpe?g|bmp|gif|tiff?|ico)$", RegexOptions.IgnoreCase) ? new[] { "toexcel" }.Concat(args).ToArray() :
                Regex.IsMatch(args[0], "^.*\\.xlsx$", RegexOptions.IgnoreCase) ? new[] { "toimage" }.Concat(args).ToArray() :
                args;

        public static void ToExcel(string[] paths)
        {
            if (paths == null || paths.Length == 0) throw new CommandArgumentsException();

            Parallel.ForEach(paths, path =>
            {
                if (!Regex.IsMatch(path, "^.*\\.(png|jpe?g|bmp|gif|tiff?|ico)$", RegexOptions.IgnoreCase)) throw new WarningException(string.Format("画像ファイルを指定してください。ファイル名: {0}", path));
                var inputPath = Path.GetFullPath(path);
                var outputPath = Path.ChangeExtension(inputPath, ".xlsx");
                if (!File.Exists(inputPath)) throw new WarningException(string.Format("ファイル {0} にアクセスできません。", inputPath));

                var size = ImageFile.GetSize(inputPath);

                ImageFile.Load(inputPath)
                    .SaveAsExcel(outputPath, size);
            });
        }

        public static void ToImage(string[] paths, string format = "png")
        {
            if (paths == null || paths.Length == 0) throw new CommandArgumentsException();
            if (!Regex.IsMatch(format, "^(png|jpe?g|bmp|gif|tiff?|ico)$", RegexOptions.IgnoreCase)) throw new CommandArgumentsException();

            Parallel.ForEach(paths, path =>
            {
                if (!Regex.IsMatch(path, "^.*\\.xlsx$", RegexOptions.IgnoreCase)) throw new WarningException(string.Format("Excel ファイル (.xlsx) を指定してください。ファイル名: {0}", path));
                var inputPath = Path.GetFullPath(path);
                var outputPath = Path.ChangeExtension(inputPath, format);
                if (!File.Exists(inputPath)) throw new WarningException(string.Format("ファイル {0} にアクセスできません。", inputPath));

                var size = ExcelFile.GetSize(inputPath);
                if (size.Width == 0 || size.Height == 0) throw new WarningException("幅または高さが 0 px です。");

                ExcelFile.Load(inputPath)
                    .SaveAsImage(outputPath, size);
            });
        }
    }
}
