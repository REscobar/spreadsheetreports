namespace SpreadSheetsReports.DocumentModel
{
    using System.Collections.Generic;

    public class Row
    {
        public Row()
        {
            this.Cells = new List<Cell>();
        }

        public IEnumerable<Cell> Cells { get; set; }

        public float? Height { get; set; }

        public CellStyle? Style { get; set; }
    }
}