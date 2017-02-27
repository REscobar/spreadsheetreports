namespace SpreadSheetsReports.ReportModel
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using SpreadSheetsReports.DocumentModel;

    public class ReportSection : DataSourceBoundReportControl, IRowCollectionGenerator
    {
        public RowCollectionSection Header { get; set; }

        public ReportSection SubSection { get; set; }

        public RowCollectionSection Footer { get; set; }

        public IEnumerable<DocumentModel.Row> Generate()
        {
            this.Databind();
            List<DocumentModel.Row> rows = new List<DocumentModel.Row>();
            if (this.Header != null)
            {
                rows.AddRange(this.Header.Generate());
            }

            if (this.SubSection != null)
            {
                if (this.DataSource != null)
                {
                    if (!string.IsNullOrWhiteSpace(this.DataMember))
                    {
                        var browser = new ObjectDataSourceBrowser(this.DataSource.GetValue(this.DataMember));
                        while (browser.MoveNext())
                        {
                            rows.AddRange(this.SubSection.Generate());
                        }
                    }
                    else
                    {
                        while (this.DataSource.MoveNext())
                        {
                            rows.AddRange(this.SubSection.Generate());
                        }
                    }
                }
                else
                {
                    rows.AddRange(this.SubSection.Generate());
                }
            }

            if (this.Footer != null)
            {
                rows.AddRange(this.Footer.Generate());
            }

            return rows;
        }
    }
}
