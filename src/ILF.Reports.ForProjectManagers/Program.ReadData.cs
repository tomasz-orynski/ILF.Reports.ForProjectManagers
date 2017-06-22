using ILF.Reports.ForProjectManagers.Diagnostics;
using ILF.Reports.ForProjectManagers.Model;
using MoreLinq;
using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ILF.Reports.ForProjectManagers
{
    partial class Program
    {
        private static ReportModel ReadReportData(string pathInputDataXlsx)
            => _logger.OnEntryCall(() =>
            {
                using (var package = new ExcelPackage(new FileInfo(pathInputDataXlsx)))
                {
                    var workbook = package.Workbook;
                    var model = new ReportModel();

                    var factory = Task.Factory;
                    var taskReadParams = factory.StartNew(() => ReadReportData_Params(model, workbook));
                    var taskReadProjects = factory.StartNew(() => ReadReportData_Projects(model, workbook));
                    var taskReadProjectTeams = taskReadProjects.ContinueWith(
                        t => ReadReportData_ProjectTeams(model, workbook),
                        TaskContinuationOptions.NotOnFaulted);
                    var taskReadProjectTeamMembers = taskReadProjectTeams.ContinueWith(
                        t => ReadReportData_ProjectTeamMembers(model, workbook),
                        TaskContinuationOptions.NotOnFaulted);

                    taskReadParams.Wait();

                    var taskRead_PMT02_MH_at_start = taskReadProjectTeamMembers
                        .ContinueWith(
                            _ => ReadReportData_PMT02_MH_at_start(model, workbook),
                            TaskContinuationOptions.NotOnFaulted);
                    var taskRead_PMT02_Cost_at_start = taskReadProjectTeamMembers
                        .ContinueWith(
                            _ => ReadReportData_PMT02_Cost_at_start(model, workbook),
                            TaskContinuationOptions.NotOnFaulted);
                    var taskRead_Utilised_MH_and_Cost_TS = taskReadProjectTeamMembers
                        .ContinueWith(
                            _ => ReadReportData_Utilised_MH_and_Cost_TS(model, workbook),
                            TaskContinuationOptions.NotOnFaulted);
                    var taskRead_Planned_Actual_Comparison = taskReadProjectTeamMembers
                        .ContinueWith(
                            _ => ReadReportData_Planned_Actual_Comparison(model, workbook),
                            TaskContinuationOptions.NotOnFaulted);
                    var taskRead_Planned_Actual_Compare_Estimate = taskReadProjectTeamMembers
                        .ContinueWith(
                            _ => ReadReportData_Planned_Actual_Compare_Estimate(model, workbook),
                            TaskContinuationOptions.NotOnFaulted);

                    Task.WaitAll(
                        taskRead_PMT02_MH_at_start,
                        taskRead_PMT02_Cost_at_start,
                        taskRead_Utilised_MH_and_Cost_TS,
                        taskRead_Planned_Actual_Comparison,
                        taskRead_Planned_Actual_Compare_Estimate);

                    return model;
                }
            });

        private static void ReadReportData_Params(ReportModel model, ExcelWorkbook workbook)
            => _logger.OnEntryCall(() =>
            {
                var sheet = workbook.Worksheets["_REPORT_"].CheckNotNull();
                model.DtStart = Convert.ToDateTime(sheet.GetValue(3, 3));
                model.DtEnd = Convert.ToDateTime(sheet.GetValue(4, 3));
            });

        private static void ReadReportData_Projects(ReportModel model, ExcelWorkbook workbook)
            => _logger.OnEntryCall(() =>
            {
                var sheet = workbook.Worksheets["Projects"].CheckNotNull();
                var row = 1;
                while (true)
                {
                    var name = sheet.GetValue<string>(++row, 1).NullTrim();
                    if (string.IsNullOrEmpty(name)) return;
                    model.Projects.Add(name);
                }
            });

        private static void ReadReportData_ProjectTeams(ReportModel model, ExcelWorkbook workbook)
            => _logger.OnEntryCall(() =>
            {
                var sheet = workbook.Worksheets["Project Teams"].CheckNotNull(); ;
                var row = 1;
                while (true)
                {
                    var name = sheet.GetValue<string>(++row, 1).NullTrim();
                    if (string.IsNullOrEmpty(name)) break;
                    model.Teams.Add(new TeamModel()
                    {
                        DivisionName = name,
                        AreaName = sheet.GetValue<string>(row, 2).NullTrim(),
                        TeamName = sheet.GetValue<string>(row, 3).NullTrim(),
                        TeamLeader = sheet.GetValue<string>(row, 4).NullTrim(),
                        DivisionLeader = sheet.GetValue<string>(row, 5).NullTrim(),
                        DivisionLeaderEmail = sheet.GetValue<string>(row, 6).NullTrim(),
                        DivisionNameShort = sheet.GetValue<string>(row, 7).NullTrim(),
                        SaveEmailPath = sheet.GetValue<string>(row, 8).NullTrim(),

                        ProjectRows = model.Projects.ToDictionary(_ => _, _ => new RowProjDataModel()),
                    });
                }

                model._RowsDivProj = model.Teams
                    .SelectMany(t => t.ProjectRows.Select(r => new { t.DivisionName, ProjectName = r.Key, Row = r.Value }))
                    .ToDictionary(_ => (_.DivisionName, _.ProjectName), _ => _.Row);
            });

        private static void ReadReportData_ProjectTeamMembers(ReportModel model, ExcelWorkbook workbook)
            => _logger.OnEntryCall(() =>
            {
                var sheet = workbook.Worksheets["Project Team Members"].CheckNotNull();
                var row = 1;
                var teamMembers = model.Teams.ToDictionary(_ => _.DivisionName, _ => _.Members);
                while (true)
                {
                    var name = sheet.GetValue<string>(++row, 1).NullTrim();
                    if (string.IsNullOrEmpty(name)) break;
                    var division = sheet.GetValue<string>(++row, 2).NullTrim();

                    teamMembers.IfExistsValue(division, members => members.Add(new TeamMemberModel()
                    {
                        Name = name,
                        ProjectRows = model.Projects.ToDictionary(_ => _, _ => new RowDataModel()),
                    }));
                }

                model._RowsDivProjEmpl = model.Teams
                    .SelectMany(t => t.Members
                        .SelectMany(m => m.ProjectRows.Select(r => new { t.DivisionName, ProjectName = r.Key, EmployeeName = m.Name, Row = r.Value })))
                    .ToDictionary(_ => (_.DivisionName, _.ProjectName, _.EmployeeName), _ => _.Row);
            });

        private static void ReadReportData_PMT02_MH_at_start(ReportModel model, ExcelWorkbook workbook)
            => _logger.OnEntryCall(() =>
            {
                var sheet = workbook.Worksheets["PMT02_MH_at start"].CheckNotNull();
                var row = 2;
                while (true)
                {
                    var employee = sheet.GetValue<string>(++row, 1).NullTrim();
                    if (string.IsNullOrEmpty(employee)) return;
                    var division = sheet.GetValue<string>(++row, 2).NullTrim();
                    var project = sheet.GetValue<string>(++row, 6).NullTrim();

                    model._RowsDivProjEmpl.IfExistsValue((division, project, employee), rowData =>
                    {
                        rowData.Hours.E += sheet.GetValue<decimal?>(row, 7) ?? 0m;
                    });
                }
            });

        private static void ReadReportData_PMT02_Cost_at_start(ReportModel model, ExcelWorkbook workbook)
            => _logger.OnEntryCall(() =>
            {
                var sheet = workbook.Worksheets["PMT02_Cost_at start"].CheckNotNull();
                var row = 2;
                while (true)
                {
                    var employee = sheet.GetValue<string>(++row, 1).NullTrim();
                    if (string.IsNullOrEmpty(employee)) return;
                    var division = sheet.GetValue<string>(++row, 2).NullTrim();
                    var project = sheet.GetValue<string>(++row, 6).NullTrim();

                    model._RowsDivProjEmpl.IfExistsValue((division, project, employee), rowData =>
                    {
                        rowData.Costs.E += sheet.GetValue<decimal?>(row, 7) ?? 0m;
                    });
                }
            });

        private static void ReadReportData_Utilised_MH_and_Cost_TS(ReportModel model, ExcelWorkbook workbook)
            => _logger.OnEntryCall(() =>
            {
                var sheet = workbook.Worksheets["Utilised MH&Cost_TS"].CheckNotNull();
                var row = 2;
                while (true)
                {
                    var employee = sheet.GetValue<string>(++row, 5).NullTrim();
                    if (string.IsNullOrEmpty(employee)) return;
                    var division = sheet.GetValue<string>(++row, 12).NullTrim();
                    var project = sheet.GetValue<string>(++row, 8).NullTrim();

                    var dt = sheet.GetValue<DateTime>(row, 1);
                    if (dt >= model.DtStart && dt <= model.DtEnd)
                        model._RowsDivProjEmpl.IfExistsValue((division, project, employee), rowData =>
                        {
                            rowData.Hours.B += sheet.GetValue<decimal?>(row, 7) ?? 0m;
                            rowData.Costs.B += sheet.GetValue<decimal?>(row, 10) ?? 0m;
                        });
                    if (dt <= model.DtEnd)
                        model._RowsDivProj.IfExistsValue((division, project), rowData =>
                        {
                            rowData.Hours.I += sheet.GetValue<decimal?>(row, 7) ?? 0m;
                            rowData.Costs.I += sheet.GetValue<decimal?>(row, 10) ?? 0m;
                        });
                }
            });

        private static void ReadReportData_Planned_Actual_Comparison(ReportModel model, ExcelWorkbook workbook)
            => _logger.OnEntryCall(() =>
            {
                var sheet = workbook.Worksheets["Planned_Actual_Comparison"].CheckNotNull();
                var row = 5;
                var dtiStart = model.GetDtiStart();
                var dtiEnd = model.GetDtiEnd();
                while (true)
                {
                    var employee = sheet.GetValue<string>(++row, 4).NullTrim();
                    if (string.IsNullOrEmpty(employee)) return;
                    var division = sheet.GetValue<string>(++row, 12).NullTrim();
                    var project = sheet.GetValue<string>(++row, 8).NullTrim();

                    var dti = sheet.GetValue<long>(row, 8);
                    var dt = sheet.GetValue<DateTime>(row, 1);
                    if (dti >= dtiStart && dti <= dtiEnd)
                        model._RowsDivProjEmpl.IfExistsValue((division, project, employee), rowData =>
                        {
                            rowData.Hours.A += sheet.GetValue<decimal?>(row, 9) ?? 0m;
                            rowData.Costs.A += sheet.GetValue<decimal?>(row, 14) ?? 0m;
                        });
                }
            });

        private static void ReadReportData_Planned_Actual_Compare_Estimate(ReportModel model, ExcelWorkbook workbook)
            => _logger.OnEntryCall(() =>
            {
                var sheet = workbook.Worksheets["Planned_Actual_Compare Estimate"].CheckNotNull();
                var row = 5;
                var dtiStart = model.GetDtiStart();
                var dtiEnd = model.GetDtiEnd();
                while (true)
                {
                    var employee = sheet.GetValue<string>(++row, 4).NullTrim();
                    if (string.IsNullOrEmpty(employee)) return;
                    var division = sheet.GetValue<string>(++row, 12).NullTrim();
                    var project = sheet.GetValue<string>(++row, 1).NullTrim();

                    var dti = sheet.GetValue<long>(row, 8);
                    var dt = sheet.GetValue<DateTime>(row, 1);
                    if (dti >= dtiStart)
                        model._RowsDivProjEmpl.IfExistsValue((division, project, employee), rowData =>
                        {
                            rowData.Hours.F += sheet.GetValue<decimal?>(row, 9) ?? 0m;
                            rowData.Costs.F += sheet.GetValue<decimal?>(row, 14) ?? 0m;
                        });
                }
            });
    }
}
