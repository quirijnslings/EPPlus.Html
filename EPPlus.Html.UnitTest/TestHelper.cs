using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Test
{
    internal static class TestHelper
    {
        private static readonly string INPUT_FILE = "test.xlsx";
        private static readonly string OUTPUT_FILE = "test_{0}.{1}";
        internal static ExcelWorksheet GetWorkSheet()
        {
            var cwd = Directory.GetCurrentDirectory();
            FileInfo file = new FileInfo(Path.Combine(cwd, INPUT_FILE));
            ExcelPackage xlPackage = new ExcelPackage(file);
            if (xlPackage.Workbook == null)
            {
                Console.WriteLine("Excel package has no workbook");
                return null;
            }
            if (xlPackage.Workbook.Worksheets == null)
            {
                Console.WriteLine("Excel package has workbook without worksheets");
                return null;
            }
            Console.WriteLine($"Workbook has {xlPackage.Workbook.Worksheets.Count} worksheets");
            if (xlPackage.Workbook.Worksheets.Count < 1)
            {
                Console.WriteLine("Excel package has workbook without worksheets");
                return null;
            }
            return xlPackage.Workbook.Worksheets[1];
        }

        internal static void CreateFile(string identifier, string extension, string contents)
        {
            var outputFileName = string.Format(OUTPUT_FILE, identifier, extension);
            var cwd = Directory.GetCurrentDirectory();
            File.WriteAllText(Path.Combine(cwd, "../../TestOutput", outputFileName), contents);
        }
    }
}
