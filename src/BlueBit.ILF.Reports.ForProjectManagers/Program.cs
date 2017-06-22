using BlueBit.ILF.Reports.ForProjectManagers.Diagnostics;
using NLog;
using System;
using System.Diagnostics;
using System.IO;

namespace BlueBit.ILF.Reports.ForProjectManagers
{
    partial class Program
    {
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();

        static void Main(string[] args)
        {
            MakeReport(args);
            if (Debugger.IsAttached)
            {
                Console.WriteLine("Press enter key...");
                Console.ReadLine();
            }
        }

        private static void MakeReport(string[] args)
            => _logger.OnEntryCall(() =>
            {
                var path = Path.GetFullPath(
                    args.Length > 0
                    ? args[0]
                    : @".\data");
                var pathInput = Path.Combine(path, "input");
                var pathInputDataXlsx = Path.Combine(pathInput, "data.xlsx");
                var pathInputTemplateXltm = Path.Combine(pathInput, "template.xltm");
                var pathOutput = Path.Combine(path, "output");

                Debug.Assert(Directory.Exists(pathInput));
                Debug.Assert(Directory.Exists(pathOutput));
                Debug.Assert(File.Exists(pathInputDataXlsx));
                Debug.Assert(File.Exists(pathInputTemplateXltm));

                var model = ReadReportData(pathInputDataXlsx);
                WriteReportData(model, pathInputTemplateXltm, pathOutput);
            });

    }
}
