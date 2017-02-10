using System;
using SpreadSheetsReports.DocumentModel;

namespace SpreadSheetsReports.ReportModel
{
    public class Sheet : ReportControl, ISheetGenerator
    {
        public ReportSection Content { get; set; }

        public string Name { get; set; }

        public DocumentModel.Sheet Generate()
        {
            DocumentModel.Sheet sheet = new DocumentModel.Sheet();
            sheet.Name = this.Name;
            sheet.Rows = this.Content.Generate();
            return sheet;
        }

        protected override void DoRender()
        {
            if (this.Content == null)
            {
                return;
            }

            this.Content.Render();
        }
    }
}
