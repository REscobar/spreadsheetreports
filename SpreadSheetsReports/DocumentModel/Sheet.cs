namespace SpreadSheetsReports.DocumentModel
{
    using System.Collections.Generic;
    using System.Linq;

    public class Sheet
    {
        public Sheet()
        {
            this.Rows = Enumerable.Empty<Row>();
        }

        public Sheet(IEnumerable<Row> rows)
        {
            this.Rows = new List<Row>(rows).AsReadOnly();
        }

        public string Name { get; set; }

        public IEnumerable<Row> Rows { get; set; }
    }
}