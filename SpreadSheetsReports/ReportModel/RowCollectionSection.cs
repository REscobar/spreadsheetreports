using System;
using SpreadSheetsReports.Renderer;

namespace SpreadSheetsReports.ReportModel
{
    public class RowCollectionSection : ReportControl
    {
        public RowCollection Rows { get; set; }

        protected override void DoRender(IReportRenderer renderer)
        {
            foreach (var row in this.Rows)
            {
                if (row == null)
                {
                    continue;
                }

                row.Databind();
            }
        }
    }
}