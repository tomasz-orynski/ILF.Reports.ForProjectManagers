using BlueBit.ILF.Reports.ForProjectManagers.Diagnostics;
using BlueBit.ILF.Reports.ForProjectManagers.Model;
using BlueBit.ILF.Reports.ForProjectManagers.Utils;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using MoreLinq;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace BlueBit.ILF.Reports.ForProjectManagers.Generators
{
    public abstract class GeneratorBase
    {
        protected const int HeaderRowsCount = 5;
        protected const int RowStart = 12;
        protected const int RowEmpty = 33;
        protected const int RowCountBetweenReports = 2;
        protected const string SheetName = "Report";

        protected static class LogicColumn
        {
            public const int First = 2;
            public const int Last = 15;
            public const int ProjNo = 2;
            public const int Employee = 3;
            public const int A = 4;
            public const int B = 5;
            public const int C = 6;
            public const int D = 7;
            public const int E = 8;
            public const int F = 9;
            public const int G = 10;
            public const int H = 11;
            public const int I = 12;
            public const int J = 13;
            public const int K = 14;
            public const int L = 15;
        }


        public Template Templates { get; set; }
        public ReportModel Report { get; set; }
        public TeamModel Team { get; set; }



        public abstract int Generate(int row);

        protected SpreadsheetDocument _document;
        protected WorkbookPart _workbookPart;
        protected Workbook _workbook;
        protected Worksheet _worksheet;
        protected SheetData _sheetData;
        protected MergeCells _mergeCells;
        //protected CalculationChain _calculationChain;

        public void SetDocument(SpreadsheetDocument document)
        {
            _document = document;
            _workbookPart = _document.WorkbookPart;
            _workbook = _workbookPart.Workbook;

            var sheet = _workbook.GetFirstChild<Sheets>()
                .Elements<Sheet>()
                .Single(_ => _.Name == SheetName);
            var worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id);
            _worksheet = worksheetPart.Worksheet;
            _sheetData = _worksheet.GetFirstChild<SheetData>().CheckNotNull();
            _mergeCells = _worksheet.Elements<MergeCells>().Single();
        }

        protected List<Row> CopyRow(int rowSrc, int rowDst, int rowCnt, bool handleMergeCells = false)
        {
            var list = new List<Row>();
            for (var rowIdx = 0; rowIdx < rowCnt; ++rowIdx)
            {
                var srcIdx = rowSrc + rowIdx;
                var dstIdx = rowDst + rowIdx;
                var dst = (Row)Templates.Rows[srcIdx].CloneNode(true); 
                dst.RowIndex.Value = (uint)dstIdx;
                dst.Elements<Cell>()
                    .ForEach(cell =>
                    {
                        var cellRef = cell.CellReference.Value.SplitSingleToRef();
                        cell.CellReference.Value = cellRef.colRef + dst.RowIndex.Value.ToString();
                    });

                _sheetData.Append(dst);
                if (handleMergeCells)
                    Templates.AddMergedCellsTo(srcIdx, dstIdx)
                        .ForEach(_ => _mergeCells.AppendChild(_));

                list.Add(dst);
            }
            return list;
        }
        private Row GetRow(int row)
            => _sheetData.Elements<Row>()
                .Single(_ => _.RowIndex.Value == row);

        private Cell GetCell(int row, string colRef)
            => GetRow(row)
            .Elements<Cell>()
            .Single(_ => _.CellReference.Value == colRef);
        private Cell GetCell(int row, int col)
            => GetCell(row, col.GetColumnRef());


        protected void SetCellValue(Row row, int col, string value)
        {
            var cell = row.Descendants<Cell>()
                .Single(_ => _.CellReference.Value.SplitSingleToRef().colRef.GetColumnIdx() == col);
            cell.DataType = CellValues.InlineString;
            cell.InlineString = new InlineString(new Text { Text = value });
        }
        protected void SetCellValue(string cellReference, string value)
        {
            var cell = _sheetData.Descendants<Cell>()
                .Single(_ => _.CellReference.Value == cellReference);
            cell.DataType = CellValues.InlineString;
            cell.InlineString = new InlineString(new Text { Text = value });
        }
        protected void SetCellValue(string cellReference, DateTime value)
        {
            var cell = _sheetData.Descendants<Cell>()
                .Single(_ => _.CellReference.Value == cellReference);
            //cell.DataType = CellValues.Date;
            cell.CellValue = new CellValue(value.ToOADate().ToString());
        }
        protected void SetCellValue(string cellReference, Decimal value)
        {
            var cell = _sheetData.Descendants<Cell>()
                .Single(_ => _.CellReference.Value == cellReference);
            cell.DataType = CellValues.Number;
            cell.CellValue = new CellValue(value.ToString());
        }
        protected void SetCellValue(Row row, int col, Decimal value)
        {
            var cell = row.Descendants<Cell>()
                .Single(_ => _.CellReference.Value.SplitSingleToRef().colRef.GetColumnIdx() == col);
            cell.DataType = CellValues.Number;
            cell.CellValue = new CellValue(value.ToString());
        }

        protected void SetCellFormula(Row row, int col, string formula)
        {
            var cell = row.Descendants<Cell>()
                .Single(_ => _.CellReference.Value.SplitSingleToRef().colRef.GetColumnIdx() == col);
            cell.DataType = CellValues.Number;
            cell.CellFormula = new CellFormula()
            {
                Text = string.Format(formula, row.RowIndex.Value),
            };
        }

    }
}
