namespace SpreadSheetsReports.ReportModel
{
    using System;
    using System.Collections.Generic;
    using Renderer;

    public class Row : ReportControl
    {
        public List<Cell> Cells { get; set; }

        public float? Height { get; set; }

        protected override void DoRender(IReportRenderer renderer)
        {
            foreach (var cell in this.Cells)
            {
                if (cell == null)
                {
                    continue;
                }

                cell.Render(renderer);
            }
        }
    }
}
