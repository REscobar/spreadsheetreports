namespace SpreadSheetsReports.ReportModel
{
    using System.Collections.Generic;

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
                    DataSourceBrowser browser = this.GetBrowser();
                    while (this.DataSource.MoveNext())
                    {
                        rows.AddRange(this.SubSection.Generate());
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

        private DataSourceBrowser GetBrowser()
        {
            DataSourceBrowser browser;
            if (!string.IsNullOrWhiteSpace(this.DataMember))
            {
                browser = new ObjectDataSourceBrowser(this.DataSource.GetValue(this.DataMember));
            }
            else
            {
                browser = this.DataSource;
            }

            return browser;
        }
    }
}
