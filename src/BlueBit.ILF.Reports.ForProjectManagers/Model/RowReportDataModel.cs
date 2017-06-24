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
        /// H: Discipline
        /// </summary>
        public decimal E { get; set; }
        /// <summary>
        /// I: Total MH
        /// </summary>
        public decimal F { get; set; }

        public bool HasValues => A != 0 || B != 0 || F != 0;

        public void Aggregate(RowReportDataModel other)
        {
            Contract.Assert(other != null);
            A += other.A;
            B += other.B;
            E += other.E;
            F += other.F;
        }
    }

    public class RowReportProjDataModel
    {
        /// <summary>
        /// L: Estimate
        /// </summary>
        public decimal I { get; set; }
    }
}
