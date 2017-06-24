using BlueBit.ILF.Reports.ForProjectManagers.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace BlueBit.ILF.Reports.ForProjectManagers.Generators
{
    public class ReportGeneratorPrepare : GeneratorBase
    {
        public override int Generate(int row)
        {
            Templates.Rows = _sheetData.Elements<Row>()
                .Where(_ => _.RowIndex.Value >= RowStart && _.RowIndex.Value < 50)
                .ToDictionary(_ => (int)_.RowIndex.Value, _ => (Row)_.CloneNode(true));

            _sheetData.Elements<Row>()
                .Where(_ => _.RowIndex.Value >= RowStart)
                .ToList()
                .ForEach(_ => _.Remove());

            SetCellValue("H2", Report.DtStart);
            SetCellValue("J2", Report.DtEnd);
            SetCellValue("C4", Team.AreaName);
            SetCellValue("C5", Team.DivisionNameShort);
            SetCellValue("C6", Team.TeamName);
            SetCellValue("C7", Team.TeamName);
            SetCellValue("C8", Team.DivisionLeader);

            return RowStart;
        }
    }

    public class ReportGeneratorFinish : GeneratorBase
    {
        public override int Generate(int row)
        {
            row += CopyRow(RowEmpty, row, 1).Count;

            var props = _workbook.CalculationProperties;
            props.ForceFullCalculation = true;
            props.FullCalculationOnLoad = true;

            _sheetData
                .SelectMany(data => data.Elements<Row>())
                .SelectMany(_ => _.Elements<Cell>())
                .Where(_ => _.CellFormula != null && _.CellValue != null)
                .ToList()
                .ForEach(_ => _.CellValue.Remove());

            _worksheet.SheetFormatProperties.OutlineLevelRow = 1;

            return row;
        }
    }

    public class ReportGeneratorSeparator : GeneratorBase
    {
        public override int Generate(int row)
        {
            row += CopyRow(RowEmpty, row, RowCountBetweenReports).Count;
            return row;
        }
    }

    public abstract class ReportGeneratorBase : GeneratorBase
    {
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


        protected abstract int TemplateRowStart { get; }
        protected abstract RowReportProjDataModel GetData(RowProjDataModel data);
        protected abstract RowReportDataModel GetData(RowDataModel data);

        public override int Generate(int row)
        {
            row += CopyHeader(row);
     
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
                    var dstRow = CopyRowData(row, projIdx == 0 && memIdx == 0 ? RowDataType.First : RowDataType.Next);
                    //SetOutlineLevel(dstRow, 2);
                    dstRow.OutlineLevel = 2;
                    dstRow.Collapsed = true;
                    dstRow.Hidden = true;
                    SetCellValue(dstRow, ColumnNo.ProjNo, projectData.Name);
                    SetCellValue(dstRow, ColumnNo.Employee, memberData.Name);
                    SetCellValue(dstRow, ColumnNo.A, memberData.Data.A);
                    SetCellValue(dstRow, ColumnNo.B, memberData.Data.B);
                    SetCellFormula(dstRow, ColumnNo.C, Formula_C);
                    SetCellFormula(dstRow, ColumnNo.D, Formula_D);
                    SetCellValue(dstRow, ColumnNo.I, memberData.Data.I);
                    projectRowSum.Aggregate(memberData.Data);
                    row++;
                }
                {
                    var dstRow = CopyRowSum(row, projIdx == data.Count - 1 ? RowSumType.Last : RowSumType.Prev);
                    //SetOutlineLevel(dstRow, 1);
                    dstRow.OutlineLevel = 1;
                    SetCellValue(dstRow, ColumnNo.ProjNo, $"Total {projectData.Name}");
                    SetCellValue(dstRow, ColumnNo.A, projectRowSum.A);
                    SetCellValue(dstRow, ColumnNo.B, projectRowSum.B);
                    SetCellFormula(dstRow, ColumnNo.C, Formula_C);
                    SetCellFormula(dstRow, ColumnNo.D, Formula_D);
                    SetCellValue(dstRow, ColumnNo.E, projectRowSum.E);
                    SetCellValue(dstRow, ColumnNo.F, projectRowData.F);
                    SetCellFormula(dstRow, ColumnNo.G, Formula_G);
                    SetCellFormula(dstRow, ColumnNo.H, Formula_H);
                    SetCellValue(dstRow, ColumnNo.I, projectRowSum.I);
                    SetCellFormula(dstRow, ColumnNo.J, Formula_J);
                    SetCellFormula(dstRow, ColumnNo.K, Formula_K);
                    SetCellFormula(dstRow, ColumnNo.L, Formula_L);
                    row++;
                }
            };
            return row;
        }

        private int CopyHeader(int row)
            => CopyRow(TemplateRowStart, row, HeaderRowsCount).Count;
        private Row CopyRowData(int row, RowDataType type)
            => CopyRow((int)(TemplateRowStart + HeaderRowsCount + type), row, 1)[0];
        private Row CopyRowSum(int row, RowSumType type)
            => CopyRow((int)(TemplateRowStart + HeaderRowsCount + 2 + type), row, 1)[0];

        private void SetOutlineLevel(Row row, int level)
            => row.SetAttribute(new DocumentFormat.OpenXml.OpenXmlAttribute("outlineLevel", string.Empty, level.ToString()));

        private static string Formula_C => "MAX(D{0}-E{0}, 0)";
        private static string Formula_D => "IF(D{0}=0,0,ABS(E{0}/D{0}))";
        private static string Formula_G => "IF(H{0}=0,\"No plan at start\",I{0}/H{0})";
        private static string Formula_H => "MAX(H{0}-I{0}, 0)";
        private static string Formula_J => "I{0}+L{0}";
        private static string Formula_K => "M{0}+H{0}";
        private static string Formula_L => "IF(H{0}=0,\"No plan at start\",M{0}/H{0})";
    }

    public class ReportGenerator4Hours : ReportGeneratorBase
    {
        protected override int TemplateRowStart => 12;
        protected override RowReportProjDataModel GetData(RowProjDataModel data) => data.Hours;
        protected override RowReportDataModel GetData(RowDataModel data) => data.Hours;
    }

    public class ReportGenerator4Costs : ReportGeneratorBase
    {
        protected override int TemplateRowStart => 22;
        protected override RowReportProjDataModel GetData(RowProjDataModel data) => data.Costs;
        protected override RowReportDataModel GetData(RowDataModel data) => data.Costs;
    }

}
