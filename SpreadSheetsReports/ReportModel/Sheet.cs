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
            this.Databind();
            DocumentModel.Sheet sheet = new DocumentModel.Sheet();
            sheet.Name = this.Name;
            sheet.Rows = this.Content?.Generate();
            return sheet;
        }
    }
}
