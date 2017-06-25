using BlueBit.ILF.Reports.ForProjectManagers.Diagnostics;
using BlueBit.ILF.Reports.ForProjectManagers.Generators;
using BlueBit.ILF.Reports.ForProjectManagers.Model;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using MoreLinq;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace BlueBit.ILF.Reports.ForProjectManagers
{
    partial class Program
    {
        private static IEnumerable<GeneratorBase> GetGenerators()
        {
            yield return new ReportGeneratorPrepare();
            yield return new ReportGenerator4Hours();
            yield return new ReportGeneratorSeparator();
            yield return new ReportGenerator4Costs();
            yield return new ReportGeneratorFinish();
        }


        private static List<(string path, string title, string info)> WriteReportData(ReportModel model, string pathInputTemplateXlsm, string pathOutput)
            => _logger.OnEntryCall(() =>
                model.Teams
                    .AsParallel()
                    .Select(team =>
                    {
                        var id = Guid.NewGuid().ToString();
                        _logger.Info($"WRITE BEG: #[{id}] for [{team.DivisionLeader}] - period [{model.DtStart} - {model.DtEnd}].");

                        var path = Path.Combine(pathOutput, $"Raport.[{id}].xlsm");
                        File.Copy(pathInputTemplateXlsm, path);

                        var row = 0;
                        using (var document = SpreadsheetDocument.Open(path, true))
                        {

                            var templates = new Template();
                            GetGenerators()
                                .ForEach(generator =>
                                {
                                    generator.Templates = templates;
                                    generator.Report = model;
                                    generator.Team = team;
                                    generator.SetDocument(document);
                                    row = generator.Generate(row);
                                });

                            document.Save();
                        }

                        _logger.Info($"WRITE END: #[{id}] - total rows #[{row}].");
                        return (
                            path,
                            $"Raport dla kierownika pionu { team.DivisionLeader } za okres { model.DtStart.ToString("yyyyMMdd") } - { model.DtEnd.ToString("yyyyMMdd") }",
                            $"W załączeniu raport za okres od { model.DtStart.ToString("yyyy-MM-dd") }  do { model.DtEnd.ToString("yyyy -MM-dd") } dla { team.DivisionLeader }."
                                + Environment.NewLine
                                + "Proszę o akceptację zestawienia przy użyciu przycisku 'confirmed' w załączonym pliku oraz wpisywanie swoich uwag."
                            );
                    })
                    .ToList()
                );


    }
}
