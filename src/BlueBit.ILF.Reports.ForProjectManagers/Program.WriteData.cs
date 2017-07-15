using BlueBit.ILF.Reports.ForProjectManagers.Diagnostics;
using BlueBit.ILF.Reports.ForProjectManagers.Generators;
using BlueBit.ILF.Reports.ForProjectManagers.Model;
using DocumentFormat.OpenXml.Packaging;
using MoreLinq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

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


        private static List<SendData> WriteReportData(ReportModel model, string pathInputTemplateXlsm, string pathInputTemplateTxt, string pathOutput)
            => _logger.OnEntryCall(() => {
                var bodyTemplate = File.ReadAllText(pathInputTemplateTxt, Encoding.UTF8);
                return model.Teams
                    .AsParallel()
                    .Select(team =>
                    {
                        var id = Guid.NewGuid().ToString();
                        _logger.Info($"WRITE BEG: #[{id}] for [{team.DivisionLeader}] - period [{model.DtStart} - {model.DtEnd}].");

                        var path = Path.Combine(pathOutput, $"Raport.({id}).xlsm");
                        var name = $"Raport dla kierownika pionu { team.DivisionLeader } - { team.TeamName } za okres { model.DtStart.ToString("yyyyMMdd") } - { model.DtEnd.ToString("yyyyMMdd") }";
                        File.Copy(pathInputTemplateXlsm, path);
                        var row = 0;
                        using (var document = SpreadsheetDocument.Open(path, true))
                        {
                            var templates = new TemplateModel()
                            {
                                Name = name,
                            };
                            GetGenerators()
                                .ForEach(generator =>
                                {
                                    generator.Template = templates;
                                    generator.Report = model;
                                    generator.Team = team;
                                    generator.SetDocument(document);
                                    row = generator.Generate(row);
                                });
                            document.Save();
                        }

                        _logger.Info($"WRITE END: #[{id}] - total rows #[{row}].");
                        return new SendData()
                        {
                            ID = id,
                            AttachmentPath = path,
                            Title = name,
                            MsgBody = bodyTemplate
                                .Replace("{model.DtStart}", model.DtStart.ToString("dd.MM.yyyy"))
                                .Replace("{model.DtEnd}", model.DtEnd.ToString("dd.MM.yyyy"))
                                .Replace("{team.DivisionLeader}", team.DivisionLeader)
                                .Replace("{team.TeamName}", team.TeamName),
                            AddressTo = team.DivisionLeaderEmail,
                        };
                    })
                    .ToList();
                });


    }
}
