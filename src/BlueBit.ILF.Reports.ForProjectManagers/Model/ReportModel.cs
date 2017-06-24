using System;
using System.Collections.Generic;

namespace BlueBit.ILF.Reports.ForProjectManagers.Model
{
    public class ReportModel
    {
        public DateTime DtStart { get; set; }
        public DateTime DtEnd { get; set; }

        public long GetDtiStart() => DtStart.Year * 100 + DtStart.Month;
        public long GetDtiEnd() => DtEnd.Year * 100 + DtEnd.Month;

        public List<string> Projects { get; set; }
        public List<TeamModel> Teams { get; } = new List<TeamModel>();

        public IReadOnlyDictionary<(string divisionName, string projectName), RowProjDataModel> _RowsDivProj { get; set; }
        public IReadOnlyDictionary<(string divisionName, string projectName, string employeeName), RowDataModel> _RowsDivProjEmpl { get; set; }
    }
}
