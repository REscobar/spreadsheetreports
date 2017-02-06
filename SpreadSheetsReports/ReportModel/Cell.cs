namespace SpreadSheetsReports.ReportModel
{
    using DocumentModel;

    public class Cell : ReportControl
    {
        public Cell()
        {
            this.Style = new CellStyle();
        }

        public CellStyle Style { get; set; }

        public string ClassName { get; set; }

        public object Value { get; set; }

        public CellType Type { get; set; }
    }
}
