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
                if (!string.IsNullOrWhiteSpace(this.DataMember))
                {
                    var browser = new ObjectDataSourceBrowser(this.DataSource.Current);
                    while (browser.MoveNext())
                    {
                        sheets.AddRange(this.Sheets.Select(s => s.Generate()));
                    }
                }
                else
                {
                    while (this.DataSource.MoveNext())
                    {
                        sheets.AddRange(this.Sheets.Select(s => s.Generate()));
                    }
                }
            }
            else
            {
                sheets.AddRange(this.Sheets.Select(s => s.Generate()));
            }

            Document doc = new Document(sheets);

            return doc;
        }
    }
}
