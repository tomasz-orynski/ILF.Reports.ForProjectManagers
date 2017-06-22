using System.Collections.Generic;

namespace BlueBit.ILF.Reports.ForProjectManagers.Model
{
    public class TeamMemberModel
    {
        public string Name { get; set; }
        public IReadOnlyDictionary<string, RowDataModel> ProjectRows { get; set; }
    }
}
