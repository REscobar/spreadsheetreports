namespace SpreadSheetsReports.DocumentModel
{
    using System.Collections.Generic;

    public class Document
    {
        public Document()
        {
            this.Sheets = new List<Sheet>();
        }

        public IEnumerable<Sheet> Sheets { get; set; }
    }
}
