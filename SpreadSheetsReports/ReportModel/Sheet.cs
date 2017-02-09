namespace SpreadSheetsReports.ReportModel
{
    using System;
    using SpreadSheetsReports.Renderer;

    public class Sheet : ReportControl
    {
        public ReportSection Content { get; set; }

        public string Name { get; set; }

        protected override void DoRender(IReportRenderer renderer)
        {
            if (this.Content == null)
            {
                return;
            }

            this.Content.Render(renderer);
        }
    }
}
