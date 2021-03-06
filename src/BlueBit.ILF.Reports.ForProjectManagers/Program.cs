﻿using BlueBit.ILF.Reports.ForProjectManagers.Diagnostics;
using BlueBit.ILF.Reports.ForProjectManagers.Model;
using NLog;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;

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
                Thread.Sleep(1000);
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
                var pathInputTemplateXlsm = Path.Combine(pathInput, "template.xlsm");
                var pathInputTemplateTxt = Path.Combine(pathInput, "template.txt");
                var pathOutput = Path.Combine(path, "output");
                var pathSend = Path.Combine(path, "send");

                if (!Directory.Exists(pathOutput))
                    Directory.CreateDirectory(pathOutput);
                if (!Directory.Exists(pathSend))
                    Directory.CreateDirectory(pathSend);

                Debug.Assert(Directory.Exists(pathInput));
                Debug.Assert(Directory.Exists(pathOutput));
                Debug.Assert(Directory.Exists(pathSend));
                Debug.Assert(File.Exists(pathInputDataXlsx));
                Debug.Assert(File.Exists(pathInputTemplateXlsm));
                Debug.Assert(File.Exists(pathInputTemplateTxt));

                Directory
                    .EnumerateFiles(pathSend)
                    .ToList()
                    .ForEach(File.Delete);

                var model = true //You can skip read from file...
                    ? ReadReportData(pathInputDataXlsx)
                    : MakeDummyReportModel();

                var result = WriteReportData(model, pathInputTemplateXlsm, pathInputTemplateTxt, pathOutput);
                SendReportData(pathSend, result);
            });


        private static ReportModel MakeDummyReportModel()
        {
            var model = new ReportModel()
            {
                DtStart = DateTime.Now,
                DtEnd = DateTime.Now,
            };
            model.Teams.Add(new TeamModel()
            {
                DivisionLeader = "leader",
                DivisionName = "name",
                SaveEmailPath = "path",
            });
            return model;
        }

    }
}
