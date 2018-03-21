using BlueBit.ILF.Reports.ForProjectManagers.Diagnostics;
using BlueBit.ILF.Reports.ForProjectManagers.Model;
using BlueBit.ILF.Reports.ForProjectManagers.Utils;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.VariantTypes;
using MoreLinq;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
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


        public TemplateModel Template { get; set; }
        public ReportModel Report { get; set; }
        public TeamModel Team { get; set; }



        public abstract int Generate(int row);

        protected SpreadsheetDocument _document;
        protected Properties _properties;
        protected WorkbookPart _workbookPart;
        protected Workbook _workbook;
        protected Worksheet _worksheet;
        protected SheetData _sheetData;
        protected MergeCells _mergeCells;

        public void SetDocument(SpreadsheetDocument document)
        {
            _document = document;
            _properties = _document.CustomFilePropertiesPart.Properties;
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
                var dst = (Row)Template.Rows[srcIdx].CloneNode(true); 
                dst.RowIndex.Value = (uint)dstIdx;
                dst.Elements<Cell>()
                    .ForEach(cell =>
                    {
                        var cellRef = cell.CellReference.Value.SplitSingleToRef();
                        cell.CellReference.Value = cellRef.colRef + dst.RowIndex.Value.ToString();
                    });

                _sheetData.Append(dst);
                if (handleMergeCells)
                    Template.AddMergedCellsTo(srcIdx, dstIdx)
                        .ForEach(_ => _mergeCells.AppendChild(_));

                list.Add(dst);
            }
            return list;
        }
        private Row GetRow(int row)
            => _sheetData.Elements<Row>()
                .Single(_ => _.RowIndex.Value == row);

        private Cell GetCell(string cellRef)
            => _sheetData.Descendants<Cell>()
                .Single(_ => _.CellReference.Value == cellRef);
        private Cell GetCell(Row row, string colRef)
            => row
                .Elements<Cell>()
                .Single(_ => _.CellReference.Value.SplitSingleToRef().colRef == colRef);

        private Cell GetCell(Row row, int col)
            => GetCell(row, col.GetColumnRef());
        private Cell GetCell(int row, string colRef)
            => GetCell(GetRow(row), colRef);
        private Cell GetCell(int row, int col)
            => GetCell(row, col.GetColumnRef());

        protected void SetCellValue(Row row, int col, string value)
            => SetCellValue(GetCell(row, col), value);
        protected void SetCellValue(string cellRef, string value)
            => SetCellValue(GetCell(cellRef), value);
        private void SetCellValue(Cell cell, string value)
        {
            cell.DataType = CellValues.InlineString;
            cell.InlineString = new InlineString(new Text { Text = value });
        }

        protected void SetCellValue(string cellRef, DateTime value)
            => SetCellValue(GetCell(cellRef), value);
        protected void SetCellValue(Cell cell, DateTime value)
        {
            cell.DataType = CellValues.Number;
            cell.CellValue = new CellValue(value.ToOADate().ToString(CultureInfo.InvariantCulture));
        }

        protected void SetCellValue(Row row, int col, Decimal value)
            => SetCellValue(GetCell(row, col), value);
        protected void SetCellValue(string cellRef, Decimal value)
            => SetCellValue(GetCell(cellRef), value);
        protected void SetCellValue(Cell cell, Decimal value)
        {
            cell.DataType = CellValues.Number;
            cell.CellValue = new CellValue(value.ToString(CultureInfo.InvariantCulture));
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

        protected void SetDocProperty(string name, string value)
        {
            var prop = _properties
                .Elements<CustomDocumentProperty>()
                .Single(_ => _.Name == name);
            prop.VTLPWSTR = new VTLPWSTR(value);
        }
    }
}
