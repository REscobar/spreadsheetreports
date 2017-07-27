namespace SpreadSheetsReports.ReportModel
{
    using System.Collections.Generic;
    using System.Linq;
    using System.Runtime.Serialization;
    using DocumentModel;

    [DataContract]
    public class ReportDefinition : DataSourceBoundReportControl, IDocumentGenerator
    {
        [DataMember]
        public List<ISheetGenerator> Sheets { get; set; }

        public Document Generate()
        {
            List<DocumentModel.Sheet> sheets = new List<DocumentModel.Sheet>();
            if (this.DataSource != null)
            {
                DataSourceBrowser browser;
                if (!string.IsNullOrWhiteSpace(this.DataMember))
                {
                    browser = new ObjectDataSourceBrowser(this.DataSource, this.DataMember);
                    while (browser.MoveNext())
                    {
                        DataBindingContext.Push(browser.Current);
                        sheets.AddRange(this.Sheets.Select(s => s.Generate()));
                        DataBindingContext.Pop();
                    }
                }
                else
                {
                    DataBindingContext.Push(this.DataSource.Current);
                    sheets.AddRange(this.Sheets.Select(s => s.Generate()));
                    DataBindingContext.Pop();
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
