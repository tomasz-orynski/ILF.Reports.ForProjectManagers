using BlueBit.ILF.Reports.ForProjectManagers.Diagnostics;
using BlueBit.ILF.Reports.ForProjectManagers.Model;
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

        protected static class ColumnNo
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


        public Templates Templates { get; set; }
        public ReportModel Report { get; set; }
        public TeamModel Team { get; set; }



        public abstract int Generate(int row);

        protected SpreadsheetDocument _document;
        protected Workbook _workbook;
        protected Worksheet _worksheet;
        protected SheetData _sheetData;

        public void SetDocument(SpreadsheetDocument document)
        {
            _document = document;
            _workbook = _document.WorkbookPart.Workbook;
            var sheet = _workbook.GetFirstChild<Sheets>()
                .Elements<Sheet>()
                .Single(_ => _.Name == SheetName);
            var worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id);
            _worksheet = worksheetPart.Worksheet;
            _sheetData = _worksheet.GetFirstChild<SheetData>().CheckNotNull();
        }

        protected List<Row> CopyRow(int rowSrc, int rowDst, int rowCnt)
        {
            var list = new List<Row>();
            for (var rowIdx = 0; rowIdx < rowCnt; ++rowIdx)
            {
                var dst = (Row)Templates.Rows[rowSrc + rowIdx].CloneNode(true); 
                dst.RowIndex.Value = (uint)(rowDst + rowIdx);
                dst.Elements<Cell>()
                    .ForEach(cell =>
                    {
                        var cellIdx = SplitCellRef(cell.CellReference.Value);
                        cell.CellReference.Value = cellIdx.col + dst.RowIndex.Value.ToString();
                    });

                _sheetData.Append(dst);
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
            => GetCell(row, GetColumnRef(col));


        protected void SetCellValue(Row row, int col, string value)
        {
            var cell = row.Descendants<Cell>()
                .Single(_ => GetColumnIndex(SplitCellRef(_.CellReference.Value).col) == col);
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
                .Single(_ => GetColumnIndex(SplitCellRef(_.CellReference.Value).col) == col);
            cell.DataType = CellValues.Number;
            cell.CellValue = new CellValue(value.ToString());
        }

        protected void SetCellFormula(Row row, int col, string formula)
        {
            var cell = row.Descendants<Cell>()
                .Single(_ => GetColumnIndex(SplitCellRef(_.CellReference.Value).col) == col);
            cell.DataType = CellValues.Number;
            cell.CellFormula = new CellFormula()
            {
                Text = string.Format(formula, row.RowIndex.Value),
            };
        }

        private static string GetA1Range(int startRow, int startColumn, int endRow, int endColumn)
            => GetColumnRef(startColumn) + startRow.ToString() + ":" + GetColumnRef(endColumn) + endRow.ToString();

        private static string GetColumnRef(int columnIndex)
        {
            int dividend = columnIndex;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }
            return columnName;
        }

        protected int GetColumnIndex(string reference)
        {
            int ci = 0;
            reference = reference.ToUpper();
            for (int ix = 0; ix < reference.Length && reference[ix] >= 'A'; ix++)
                ci = (ci * 26) + ((int)reference[ix] - 64);
            return ci;
        }

        protected (string row, string col) SplitCellRef(string reference)
            => (string.Concat(reference.Where(char.IsDigit)), string.Concat(reference.Where(char.IsLetter)));

        protected (int row, int col) GetCellIndex(string reference)
        {
            var splitedRef = SplitCellRef(reference);
            return (Convert.ToInt32(splitedRef.row), GetColumnIndex(splitedRef.col));
        }

        protected string GetCellRef(int row, int col)
            => GetColumnRef(col) + row.ToString();

    }
}
