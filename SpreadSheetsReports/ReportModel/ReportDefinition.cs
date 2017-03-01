namespace SpreadSheetsReports.ReportModel
{
    using System.Collections.Generic;
    using System.Linq;
    using DocumentModel;

    public class ReportDefinition : DataSourceBoundReportControl, IDocumentGenerator
    {
        public List<Sheet> Sheets { get; set; }

        public Document Generate()
        {
            List<DocumentModel.Sheet> sheets = new List<DocumentModel.Sheet>();
            if (this.DataSource != null)
            {
                DataSourceBrowser browser = this.GetDataBrowser();

                while (this.DataSource.MoveNext())
                {
                    sheets.AddRange(this.Sheets.Select(s => s.Generate()));
                }
            }
            else
            {
                sheets.AddRange(this.Sheets.Select(s => s.Generate()));
            }

            Document doc = new Document(sheets);

            return doc;
        }

        private DataSourceBrowser GetDataBrowser()
        {
            DataSourceBrowser browser;
            if (!string.IsNullOrWhiteSpace(this.DataMember))
            {
                browser = new ObjectDataSourceBrowser(this.DataSource, this.DataMember);
            }
            else
            {
                browser = this.DataSource;
            }

            return browser;
        }
    }
}
