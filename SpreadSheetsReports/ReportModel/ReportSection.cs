using System;
using SpreadSheetsReports.Renderer;

namespace SpreadSheetsReports.ReportModel
{
    public class ReportSection : ReportControl
    {
        public RowCollectionSection Header { get; set; }

        public ReportSection SubSection { get; set; }

        public RowCollectionSection Footer { get; set; }

        protected override void DoRender(IReportRenderer renderer)
        {
            if (this.Header != null)
            {
                this.Header.Render(renderer);
            }

            if (this.SubSection != null)
            {
                this.SubSection.Render(renderer);
            }

            if (this.SubSection != null)
            {
                this.Footer.Render(renderer);
            }
        }
    }
}
