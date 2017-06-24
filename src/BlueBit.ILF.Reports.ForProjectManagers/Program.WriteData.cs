using BlueBit.ILF.Reports.ForProjectManagers.Diagnostics;
using BlueBit.ILF.Reports.ForProjectManagers.Model;
using MoreLinq;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using System.Linq;

namespace BlueBit.ILF.Reports.ForProjectManagers
{
    partial class Program
    {
        private abstract class ReportGenerator
        {
            const int HeaderRowsCount = 5;
            static class Column 
            {
                public const int First = 2;
                public const int Last = 15;
                public const int ProjNo = 2;
                public const int Employee = 3;
                public const int A = 4;
                public const int B = 5;
                public const int E = 8;
                public const int F = 9;
                public const int I = 12;
            }
            enum RowDataType : int
            {
                First = 0,
                Next = 1,
            }
            enum RowSumType : int
            {
                Prev = 0,
                Last = 1,
            }

            public ReportModel Report { get; set; }
            public TeamModel Team { get; set; }

            public ExcelWorksheet Sheet { get; set; }

            protected abstract int TemplateRowStart { get; }
            protected abstract RowReportProjDataModel GetData(RowProjDataModel data);
            protected abstract RowReportDataModel GetData(RowDataModel data);

            public int Generate(int row)
            {
                row = CopyHeader(row);
                Sheet.SetValue("H2", Report.DtStart);
                Sheet.SetValue("J2", Report.DtEnd);
                Sheet.SetValue("C4", Team.AreaName);
                Sheet.SetValue("C5", Team.DivisionNameShort);
                Sheet.SetValue("C6", Team.TeamName);
                Sheet.SetValue("C7", Team.TeamName);
                Sheet.SetValue("C8", Team.DivisionLeader);

                var data = Report.Projects
                    .Select(project => new {
                        Name = project,
                        MembersData = Team.Members
                            .Where(_ => _.ProjectRows.ContainsKey(project))
                            .OrderBy(_ => _.Name)
                            .Select(_ => new { _.Name, Data = GetData(_.ProjectRows[project]) })
                            .Where(_ => _.Data.HasValues)
                            .ToList()
                    })
                    .Where(_ => _.MembersData.Count > 0)
                    .ToList();

                for (var projIdx = 0; projIdx < data.Count; ++projIdx)
                {
                    var projectData = data[projIdx];
                    var isLastProject = projIdx == data.Count - 1;
                    var projectRowData = GetData(Team.ProjectRows[projectData.Name]);
                    var projectRowSum = new RowReportDataModel();

                    for (var memIdx = 0; memIdx < projectData.MembersData.Count; ++memIdx)
                    {
                        var memberData = projectData.MembersData[memIdx];
                        CopyRowData(row, projIdx == 0 && memIdx == 0 ? RowDataType.First : RowDataType.Next);
                        Sheet.SetValue(row, Column.ProjNo, projectData.Name);
                        Sheet.SetValue(row, Column.Employee, memberData.Name);
                        Sheet.SetValue(row, Column.A, memberData.Data.A);
                        Sheet.SetValue(row, Column.B, memberData.Data.B);
                        Sheet.SetValue(row, Column.I, projectRowData.I);
                        projectRowSum.Aggregate(memberData.Data);
                        row++;
                    }
                    {
                        CopyRowSum(row, projIdx == data.Count - 1 ? RowSumType.Last : RowSumType.Prev);
                        Sheet.SetValue(row, Column.ProjNo, $"Total {projectData.Name}");
                        Sheet.SetValue(row, Column.A, projectRowSum.A);
                        Sheet.SetValue(row, Column.B, projectRowSum.B);
                        Sheet.SetValue(row, Column.E, projectRowSum.E);
                        Sheet.SetValue(row, Column.F, projectRowSum.F);
                        Sheet.SetValue(row, Column.I, projectRowData.I);
                        row++;
                    }
                };
                return row;
            }

            private int CopyRow(int rowSrc, int rowDst, int rowCnt)
            {
                for (var rowIdx = 0; rowIdx < rowCnt; ++rowIdx)
                    for (var col = Column.First; col <= Column.Last; ++col)
                    {
                        Sheet.SetValue(rowDst + rowIdx, col, Sheet.GetValue(rowSrc + rowIdx, col));
                        var rngSrc = Sheet.Cells[rowSrc + rowIdx, col];
                        var rngDst = Sheet.Cells[rowDst + rowIdx, col];
                        rngDst.StyleID = rngSrc.StyleID;
                    }
                return rowDst + rowCnt;
            }

            private int CopyHeader(int row)
                => CopyRow(TemplateRowStart, row, HeaderRowsCount);
            private void CopyRowData(int row, RowDataType type)
                => CopyRow((int)(TemplateRowStart + HeaderRowsCount + type), row, 1);
            private void CopyRowSum(int row, RowSumType type)
                => CopyRow((int)(TemplateRowStart + HeaderRowsCount + 2 + type), row, 1);
        }

        private class ReportGenerator4Hours : ReportGenerator
        {
            protected override int TemplateRowStart => 12;
            protected override RowReportProjDataModel GetData(RowProjDataModel data) => data.Hours;
            protected override RowReportDataModel GetData(RowDataModel data) => data.Hours;
        }

        private class ReportGenerator4Costs : ReportGenerator
        {
            protected override int TemplateRowStart => 22;
            protected override RowReportProjDataModel GetData(RowProjDataModel data) => data.Costs;
            protected override RowReportDataModel GetData(RowDataModel data) => data.Costs;
        }

        private static IEnumerable<ReportGenerator> GetGenerators()
        {
            yield return new ReportGenerator4Hours();
            yield return new ReportGenerator4Costs();
        }

        const int RowStart = 35;
        const int RowCountBetweenReports = 2;

        private static List<(string path, string title, string info)> WriteReportData(ReportModel model, string pathInputTemplateXlsm, string pathOutput)
            => _logger.OnEntryCall(() =>
                model.Teams
                    .AsParallel()
                    .Select(team =>
                    {
                        var dt = DateTime.Now.ToString("yyyyMMddhhmmss");
                        var path = Path.Combine(pathOutput, $"Raport.{ dt }.xlsm");
                        File.Copy(pathInputTemplateXlsm, path);

                        using (var package = new ExcelPackage(new FileInfo(path)))
                        using (var workbook = package.Workbook)
                        using (var sheet = workbook.Worksheets["Report"].CheckNotNull())
                        {
                            workbook.Properties.SetCustomPropertyValue("_SAVE_PATH_", team.SaveEmailPath);
                            var row = RowStart;
                            GetGenerators()
                                .ForEach(generator =>
                                {
                                    generator.Report = model;
                                    generator.Team = team;
                                    generator.Sheet = sheet;
                                    row = generator.Generate(row) + RowCountBetweenReports;
                                });
                            sheet.DeleteRow(12, RowStart - 12);
                            workbook.Styles.UpdateXml();
                            package.Save();
                        }
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
