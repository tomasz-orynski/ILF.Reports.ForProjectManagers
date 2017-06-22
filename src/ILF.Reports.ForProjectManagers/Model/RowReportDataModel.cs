namespace ILF.Reports.ForProjectManagers.Model
{
    public class RowReportDataModel
    {
        /// <summary>
        /// D: Planned
        /// </summary>
        public decimal A { get; set; }
        /// <summary>
        /// E: Consumed
        /// </summary>
        public decimal B { get; set; }
        /// <summary>
        /// H: Discipline
        /// </summary>
        public decimal E { get; set; }
        /// <summary>
        /// I: Total MH
        /// </summary>
        public decimal F { get; set; }
    }

    public class RowReportProjDataModel
    {
        /// <summary>
        /// L: Estimate
        /// </summary>
        public decimal I { get; set; }
    }
}
