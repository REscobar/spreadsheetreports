namespace SpreadSheetsReports.ReportModel
{
    using System.Collections.Generic;

    public class Row : ReportControl
    {
        public List<Cell> Cells { get; set; }

        public float? Height { get; set; }

        internal override void DoRender()
        {
            base.DoRender();

            foreach (var cell in this.Cells)
            {
                if (cell == null)
                {
                    continue;
                }

                cell.DoRender();
            }
        }
    }
}
