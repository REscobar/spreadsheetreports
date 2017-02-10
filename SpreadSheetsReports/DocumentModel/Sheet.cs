namespace SpreadSheetsReports.DocumentModel
{
    using System.Collections.Generic;

    public class Sheet
    {
        public Sheet()
        {
            this.Rows = new List<Row>();
        }

        public string Name { get; internal set; }
        public IEnumerable<Row> Rows { get; set; }
    }
}