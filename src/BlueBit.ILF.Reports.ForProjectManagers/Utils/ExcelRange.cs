using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.Linq;

namespace BlueBit.ILF.Reports.ForProjectManagers.Utils
{
    public static class ExcelRange
    {
        public static string GetColumnRef(this int columnIdx)
        {
            string columnName = String.Empty;
            int modulo;
            while (columnIdx > 0)
            {
                modulo = (columnIdx - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                columnIdx = (int)((columnIdx - modulo) / 26);
            }
            return columnName;
        }

        public static int GetColumnIdx(this string columnRef)
        {
            int ci = 0;
            columnRef = columnRef.ToUpper();
            for (int ix = 0; ix < columnRef.Length && columnRef[ix] >= 'A'; ix++)
                ci = (ci * 26) + ((int)columnRef[ix] - 64);
            return ci;
        }

        public static (string rowRef, string colRef) SplitSingleToRef(this string reference)
            => (string.Concat(reference.Where(char.IsDigit)), string.Concat(reference.Where(char.IsLetter)));

        public static (int rowIdx, int colIdx) SplitSingleToIdx(this string reference)
        {
            var splitedRef = SplitSingleToRef(reference);
            return (Convert.ToInt32(splitedRef.rowRef), GetColumnIdx(splitedRef.colRef));
        }

        public static string MergeToRef(int rowIdx, int colIdx)
            => GetColumnRef(colIdx) + rowIdx.ToString();

        public static string MergeToRef(int startRow, int startColumn, int endRow, int endColumn)
            => $"{GetColumnRef(startColumn)}{startRow}:{GetColumnRef(endColumn)}{endRow}";
        public static string MergeToRef(int startRow, string startColumn, int endRow, string endColumn)
            => $"{startColumn}{startRow}:{endColumn}{endRow}";

        public static string MergeToRef(this IEnumerable<string> references)
            => string.Join(",", references);


        public static IEnumerable<int> GetRows(this string range)
        {
            if (range.Contains(":"))
            {
                var rangeSplited = range.Split(':')
                    .Select(_ => _.SplitSingleToIdx().rowIdx)
                    .OrderBy(_ => _)
                    .ToList();
                var first = rangeSplited[0];
                var last = rangeSplited[1];
                for (var rowIdx = first; rowIdx <= last; ++rowIdx)
                    yield return rowIdx;
            }
            else
            {
                yield return range.SplitSingleToIdx().rowIdx;
            }
        }

        public static (string firstCol, string lastCol) GetCols(this string range)
        {
            Contract.Assert(range.Contains(":"));
            var rangeSplited = range.Split(':')
                .Select(_ => _.SplitSingleToRef().colRef)
                .OrderBy(_ => _)
                .ToList();
            return (rangeSplited[0], rangeSplited[1]);
        }
    }
}
