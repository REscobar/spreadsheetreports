namespace SpreadSheetsReports.ReportModel
{
    using System.Collections.Generic;

    public class ReportSection : DataSourceBoundReportControl, IRowCollectionGenerator
    {
        public IRowCollectionGenerator Header { get; set; }

        public IRowCollectionGenerator SubSection { get; set; }

        public IRowCollectionGenerator Footer { get; set; }

        public IEnumerable<DocumentModel.Row> Generate()
        {
            this.Databind();
            List<DocumentModel.Row> rows = new List<DocumentModel.Row>();

            DataSourceBrowser browser;
            if (!string.IsNullOrWhiteSpace(this.DataMember))
            {
                browser = new ObjectDataSourceBrowser((this.DataSource ?? new ObjectDataSourceBrowser(DataBindingContext.Peek())).GetValue(this.DataMember));
            }
            else
            {
                browser = this.DataSource;
            }

            if (browser != null)
            {
                while (browser.MoveNext())
                {
                    DataBindingContext.Push(browser.Current);

                    if (this.Header != null)
                    {
                        rows.AddRange(this.Header.Generate());
                    }

                    if (this.SubSection != null)
                    {
                        rows.AddRange(this.SubSection.Generate());
                    }

                    if (this.Footer != null)
                    {
                        rows.AddRange(this.Footer.Generate());
                    }
                }

                DataBindingContext.Pop();
            }
            else
            {
                if (this.Header != null)
                {
                    rows.AddRange(this.Header.Generate());
                }

                if (this.SubSection != null)
                {
                    rows.AddRange(this.SubSection.Generate());
                }

                if (this.Footer != null)
                {
                    rows.AddRange(this.Footer.Generate());
                }
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
