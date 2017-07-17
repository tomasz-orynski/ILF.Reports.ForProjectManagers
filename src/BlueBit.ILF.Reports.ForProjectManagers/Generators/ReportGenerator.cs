using BlueBit.ILF.Reports.ForProjectManagers.Model;
using BlueBit.ILF.Reports.ForProjectManagers.Utils;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.VariantTypes;
using MoreLinq;
using System.Linq;

namespace BlueBit.ILF.Reports.ForProjectManagers.Generators
{
    public class ReportGeneratorPrepare : GeneratorBase
    {
        public override int Generate(int row)
        {
            var rows = _sheetData.Elements<Row>()
                .Where(_ => _.RowIndex.Value >= RowStart)
                .ToDictionary(_ => (int)_.RowIndex.Value, _ => _);

            var mergedCells = _mergeCells.Cast<MergeCell>()
                .Select(_ => new {
                    MergeCell = _,
                    Rows = _.Reference.Value.GetRows().ToList()
                })
                .Where(_ => _.Rows.Count == 1)
                .Select(_ => new
                {
                    _.MergeCell,
                    Row = _.Rows[0],
                })
                .Where(_ => _.Row >= RowStart)
                .GroupBy(_ => _.Row, _ => _.MergeCell)
                .ToDictionary(_ => _.Key, _ => _.ToList());


            Template.ConditionalFormatings = _worksheet.Elements<ConditionalFormatting>()
                .Select(_ => new
                {
                    ConditionalFormatting = _,
                    Col = _.SequenceOfReferences.Items.First().Value.SplitSingleToIdx().colIdx
                })
                .GroupBy(_ => _.Col, _=> _.ConditionalFormatting)
                .ToDictionary(_ => _.Key, _ => _.ToList());
            Template.ConditionalFormatings
                .SelectMany(_ => _.Value)
                .ForEach(_ => _.SequenceOfReferences.Items.Clear());

            Template.Rows = rows
                .Where(_ => _.Key < 50)
                .ToDictionary(_ => _.Key, _ => (Row)_.Value.CloneNode(true));

            Template.MergeCells = mergedCells
                .Where(_ => _.Key >= RowStart)
                .ToDictionary(
                    _ => _.Key, 
                    _ => _.Value
                        .Select(item => item.Reference.Value.GetCols())
                        .ToList()
                    );

            rows.ForEach(_ => _.Value.Remove());
            mergedCells.ForEach(_ => _.Value.ForEach(x => x.Remove()));
            _workbookPart.DeletePart(_workbookPart.CalculationChainPart);

            SetCellValue("H2", Report.DtStart);
            SetCellValue("J2", Report.DtEnd);
            SetCellValue("C4", Team.AreaName);
            SetCellValue("C5", Team.DivisionNameShort);
            SetCellValue("C6", Team.TeamName);
            SetCellValue("C7", Team.TeamLeader);
            SetCellValue("C8", Team.DivisionLeader);
            SetDocProperty("_SAVE_PATH_", Team.SaveEmailPath);
            SetDocProperty("_SAVE_NAME_", Template.Name);
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
            Mid = 1,
            Last = 2,

            _Skip = Last + 1,
        }
        enum RowSumType : int
        {
            Mid = 0,
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

            if (data.Count == 0) return row;

            var rowStart = row;
            for (var projIdx = 0; projIdx < data.Count; ++projIdx)
            {
                var projectData = data[projIdx];
                var isLastProject = projIdx == data.Count - 1;
                var projectRowData = GetData(Team.ProjectRows[projectData.Name]);
                var projectRowSum = new RowReportDataModel();

                for (var memIdx = 0; memIdx < projectData.MembersData.Count; ++memIdx)
                {
                    var memberData = projectData.MembersData[memIdx];
                    var rowType = RowDataType.Mid;
                    if (memIdx == 0)
                        rowType = RowDataType.First;
                    else if (memIdx == projectData.MembersData.Count - 1)
                        rowType = RowDataType.Last;

                    var dstRow = CopyRowData(row, rowType);
                    //SetOutlineLevel(dstRow, 2);
                    dstRow.OutlineLevel = 2;
                    dstRow.Collapsed = true;
                    dstRow.Hidden = true;
                    SetCellValue(dstRow, LogicColumn.ProjNo, projectData.Name);
                    SetCellValue(dstRow, LogicColumn.Employee, memberData.Name);
                    SetCellValue(dstRow, LogicColumn.A, memberData.Data.A);
                    SetCellValue(dstRow, LogicColumn.B, memberData.Data.B);
                    SetCellFormula(dstRow, LogicColumn.C, Formula_C);
                    SetCellFormula(dstRow, LogicColumn.D, Formula_D);
                    SetCellValue(dstRow, LogicColumn.I, memberData.Data.I);
                    projectRowSum.Aggregate(memberData.Data);
                    row++;
                }

                {
                    var dstRow = CopyRowSum(row, projIdx == data.Count - 1 ? RowSumType.Last : RowSumType.Mid);
                    //SetOutlineLevel(dstRow, 1);
                    dstRow.OutlineLevel = 1;
                    SetCellValue(dstRow, LogicColumn.ProjNo, $"Total {projectData.Name}");
                    SetCellValue(dstRow, LogicColumn.A, projectRowSum.A);
                    SetCellValue(dstRow, LogicColumn.B, projectRowSum.B);
                    SetCellFormula(dstRow, LogicColumn.C, Formula_C);
                    SetCellFormula(dstRow, LogicColumn.D, Formula_D);
                    SetCellValue(dstRow, LogicColumn.E, projectRowData.E);
                    SetCellValue(dstRow, LogicColumn.F, projectRowData.F);
                    SetCellFormula(dstRow, LogicColumn.G, Formula_G);
                    SetCellFormula(dstRow, LogicColumn.H, Formula_H);
                    SetCellValue(dstRow, LogicColumn.I, projectRowSum.I);
                    SetCellFormula(dstRow, LogicColumn.J, Formula_J);
                    SetCellFormula(dstRow, LogicColumn.K, Formula_K);
                    SetCellFormula(dstRow, LogicColumn.L, Formula_L);
                    Template.AddConditionalFormatingTo(row, LogicColumn.H);
                    Template.AddConditionalFormatingTo(row, LogicColumn.K);
                    row++;
                }
            }
            Template.AddConditionalFormatingTo(rowStart, row - 1, LogicColumn.C);

            return row;
        }

        private int CopyHeader(int row)
            => CopyRow(TemplateRowStart, row, HeaderRowsCount, true).Count;
        private Row CopyRowData(int row, RowDataType type)
            => CopyRow((int)(TemplateRowStart + HeaderRowsCount + type), row, 1)[0];
        private Row CopyRowSum(int row, RowSumType type)
            => CopyRow(TemplateRowStart + HeaderRowsCount + (int)RowDataType._Skip + (int)type, row, 1)[0];

        private static string Formula_C => "MAX(D{0}-E{0}, 0)";
        private static string Formula_D => "IF(D{0}=0,0,ABS(E{0}/D{0}))";
        private static string Formula_G => "IF(H{0}=0,\"No plan at start\",I{0}/H{0})";
        private static string Formula_H => "MAX(H{0}-I{0}, 0)";
        private static string Formula_J => "I{0}+L{0}";
        private static string Formula_K => "M{0}-H{0}";
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
        protected override int TemplateRowStart => 23;
        protected override RowReportProjDataModel GetData(RowProjDataModel data) => data.Costs;
        protected override RowReportDataModel GetData(RowDataModel data) => data.Costs;
    }

}
