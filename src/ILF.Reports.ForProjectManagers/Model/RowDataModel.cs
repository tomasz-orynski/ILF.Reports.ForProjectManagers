namespace ILF.Reports.ForProjectManagers.Model
{
    public class RowDataModel
    {
        public RowReportDataModel Hours { get; } = new RowReportDataModel();
        public RowReportDataModel Costs { get; } = new RowReportDataModel();
    }
    public class RowProjDataModel
    {
        public RowReportProjDataModel Hours { get; } = new RowReportProjDataModel();
        public RowReportProjDataModel Costs { get; } = new RowReportProjDataModel();
    }
}
