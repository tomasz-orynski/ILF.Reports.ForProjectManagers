using System.Collections.Generic;

namespace BlueBit.ILF.Reports.ForProjectManagers.Model
{
    public class TeamModel
    {
        public string DivisionName { get; set; }
        public string DivisionLeader { get; set; }
        public string DivisionLeaderEmail { get; set; }
        public string TeamName { get; set; }
        public string TeamLeader { get; set; }
        public string AreaName { get; set; }
        public string DivisionNameShort { get; set; }
        public string SaveEmailPath { get; set; }

        public List<TeamMemberModel> Members { get; } = new List<TeamMemberModel>();

        public Dictionary<string, RowProjDataModel> ProjectRows { get; set; }
    }
}
