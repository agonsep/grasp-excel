using System;
using System.IO;

namespace HtmlToXlsx;

class Program
{
    static int Main(string[] args)
    {
        if (args.Length < 1)
        {
            Console.WriteLine("Usage: HtmlToXlsx <source-folder>");
            Console.WriteLine();
            Console.WriteLine("Converts all HTML-based .xls and .mht files in the folder to proper .xlsx files.");
            return 1;
        }

        var sourcePath = args[0];

        if (!Directory.Exists(sourcePath))
        {
            Console.Error.WriteLine($"Error: Directory not found: {sourcePath}");
            return 1;
        }

        var extensions = new[] { "*.xls", "*.mht" };
        var files = extensions.SelectMany(ext => Directory.GetFiles(sourcePath, ext)).ToArray();

        if (files.Length == 0)
        {
            Console.WriteLine("No .xls or .mht files found in the specified folder.");
            return 0;
        }

        var converter = new HtmlExcelConverter();
        int success = 0, failure = 0;

        foreach (var inputFile in files)
        {
            var outputFile = Path.Combine(
                Path.GetDirectoryName(inputFile)!,
                Path.GetFileNameWithoutExtension(inputFile) + ".xlsx");

            try
            {
                converter.Convert(inputFile, outputFile);
                Console.WriteLine($"Converted: {Path.GetFileName(inputFile)} -> {Path.GetFileName(outputFile)}");
                success++;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error converting {Path.GetFileName(inputFile)}: {ex.Message}");
                failure++;
            }
        }

        Console.WriteLine();
        Console.WriteLine($"Done. {success} succeeded, {failure} failed.");
        return failure > 0 ? 1 : 0;
    }
}
