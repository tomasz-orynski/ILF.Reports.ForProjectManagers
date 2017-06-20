using MoreLinq;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ILF.Reports.ForProjectManagers
{
    class Program
    {
        static void Main(string[] args)
        {
            var path = Path.GetFullPath(
                args.Length > 0
                ? args[0]
                : @".\data");
            var pathInput = Path.Combine(path, "input");
            var pathInputDataXlsx = Path.Combine(pathInput, "data.xlsx");
            var pathOutput = Path.Combine(path, "output");

            Debug.Assert(Directory.Exists(pathInput));
            Debug.Assert(Directory.Exists(pathOutput));
            Debug.Assert(File.Exists(pathInputDataXlsx));

            using (var package = new ExcelPackage(new FileInfo(pathInputDataXlsx)))
            {
                var workbook = package.Workbook;
                workbook.Worksheets.ForEach(worksheet =>
                {
                    Console.WriteLine(worksheet.Name);
                });
            }


            if (Debugger.IsAttached)
                Console.ReadLine();
        }
    }
}
