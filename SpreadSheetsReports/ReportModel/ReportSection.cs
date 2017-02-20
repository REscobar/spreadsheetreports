using System;
using System.Collections.Generic;
using SpreadSheetsReports.DocumentModel;
using System.Linq;

namespace SpreadSheetsReports.ReportModel
{
    public class ReportSection : ReportControl, IRowCollectionGenerator
    {
        public RowCollectionSection Header { get; set; }

        public ReportSection SubSection { get; set; }

        public RowCollectionSection Footer { get; set; }

        public IEnumerable<DocumentModel.Row> Generate()
        {
            this.Databind();
            var rows = Enumerable.Empty<DocumentModel.Row>();
            if (this.Header != null)
            {
                rows = rows.Concat(this.Header.Generate());
            }

            if (this.SubSection != null)
            {
                if (this.DataSource != null)
                {
                    if (!string.IsNullOrWhiteSpace(this.DataMember))
                    {
                        var browser = new ObjectDataSourceBrowser(this.DataSource.Current);
                        while (browser.MoveNext())
                        {
                            rows = rows.Concat(this.SubSection.Generate());
                        }
                    }
                    else
                    {
                        while (this.DataSource.MoveNext())
                        {
                            rows = rows.Concat(this.SubSection.Generate());
                        }
                    }
                }
                else
                {
                    rows = rows.Concat(this.SubSection.Generate());
                }
            }

            if (this.Footer != null)
            {
                rows = rows.Concat(this.Footer.Generate());
            }

            return rows;
        }
    }
}
