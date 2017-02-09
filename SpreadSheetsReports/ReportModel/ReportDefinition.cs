namespace SpreadSheetsReports.ReportModel
{
    using Renderer;
    using System.Collections.Generic;
    using System;

    public class ReportDefinition : ReportControl
    {
        public List<Sheet> Sheets { get; set; }

        public object DataSource { get; set; }

        protected override void DoRender(IReportRenderer renderer)
        {
            foreach (var sheet in this.Sheets)
            {
                if (sheet == null)
                {
                    continue;
                }

                sheet.Render(renderer);
            }
        }
    }
}
