using System.Diagnostics.Contracts;

namespace BlueBit.ILF.Reports.ForProjectManagers.Model
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
        /// L: Estimate
        /// </summary>
        public decimal I { get; set; }

        public bool HasValues => A != 0 || B != 0 || I != 0;

        public void Aggregate(RowReportDataModel other)
        {
            Contract.Assert(other != null);
            A += other.A;
            B += other.B;
            I += other.I;
        }
    }

    public class RowReportProjDataModel
    {
        /// <summary>
        /// H: Discipline Total Budget
        /// </summary>
        public decimal E { get; set; }
        /// <summary>
        /// I: Total MH
        /// </summary>
        public decimal F { get; set; }
    }
}
