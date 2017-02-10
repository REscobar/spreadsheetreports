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

        protected override void DoRender()
        {
            if (this.Header != null)
            {
                this.Header.Render();
            }

            if (this.SubSection != null)
            {
                this.SubSection.Render();
            }

            if (this.SubSection != null)
            {
                this.Footer.Render();
            }
        }

        public IEnumerable<DocumentModel.Row> Generate()
        {
            var rows = Enumerable.Empty<DocumentModel.Row>();
            if (this.Header != null)
            {
                rows = rows.Concat(this.Header.Generate());
            }

            if (this.SubSection != null)
            {
                rows = rows.Concat(this.SubSection.Generate());
            }

            if (this.SubSection != null)
            {
                rows = rows.Concat(this.Footer.Generate());
            }

            return rows;
        }
    }
}
