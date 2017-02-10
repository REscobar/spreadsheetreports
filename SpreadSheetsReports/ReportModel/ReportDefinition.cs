namespace SpreadSheetsReports.ReportModel
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using DocumentModel;

    public class ReportDefinition : ReportControl, IDocumentGenerator
    {
        public List<Sheet> Sheets { get; set; }

        public object DataSource { get; set; }

        public Document Generate()
        {
            Document doc = new Document();
            this.Render();

            doc.Sheets = this.Sheets.Select(s => s.Generate());
            return doc;
        }

        protected override void DoRender()
        {
            foreach (var sheet in this.Sheets)
            {
                if (sheet == null)
                {
                    continue;
                }

                sheet.Render();
            }
        }

    }
}
