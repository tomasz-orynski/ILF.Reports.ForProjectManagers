using BlueBit.ILF.Reports.ForProjectManagers.Utils;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using MoreLinq;
using System.Collections.Generic;

namespace BlueBit.ILF.Reports.ForProjectManagers.Generators
{
    public class Template
    {
        public Dictionary<int, Row> Rows;
        public Dictionary<int, List<(string firstCol, string lastCol)>> MergeCells;
        public Dictionary<int, List<ConditionalFormatting>> ConditionalFormatings;

        public void AddConditionalFormatingTo(int row, int col)
        {
            ConditionalFormatings[col]
                .ForEach(_ =>
                {
                    _.SequenceOfReferences.Items.Add(new StringValue()
                    {
                        Value = ExcelRange.MergeToRef(row, col)
                    });
                });
        }
        public void AddConditionalFormatingTo(int rowBeg, int rowEnd, int col)
        {
            ConditionalFormatings[col]
                .ForEach(_ =>
                {
                    _.SequenceOfReferences.Items.Add(new StringValue()
                    {
                        Value = ExcelRange.MergeToRef(rowBeg, col, rowEnd, col)
                    });
                });
        }

        public IEnumerable<MergeCell> AddMergedCellsTo(int rowSrc, int rowDst)
        {
            if (MergeCells.TryGetValue(rowSrc, out var mergeCells))
                foreach (var mergeCell in mergeCells)
                {
                    yield return new MergeCell()
                    {
                        Reference = new StringValue()
                        {
                            Value = ExcelRange.MergeToRef(rowDst, mergeCell.firstCol, rowDst, mergeCell.lastCol)
                        }
                    };
                }
        }
    }
}
