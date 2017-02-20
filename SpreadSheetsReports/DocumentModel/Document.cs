namespace SpreadSheetsReports.DocumentModel
{
    using System.Collections.Generic;
    using System.Linq;

    public class Document
    {
        public Document()
        {
            this.Sheets = Enumerable.Empty<Sheet>();
        }

        public Document(IEnumerable<Sheet> sheets)
        {
            this.Sheets = new List<Sheet>(sheets).AsReadOnly();
        }

        public IEnumerable<Sheet> Sheets { get; set; }
    }
}
