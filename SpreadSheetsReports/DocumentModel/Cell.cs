namespace SpreadSheetsReports.DocumentModel
{
    using SpreadSheetsReports.ReportModel;

    public class Cell
    {
        public CellStyle Style { get; set; }

        public string ClassName { get; set; }

        public object Value { get; set; }

        public CellType Type { get; set; }
    }
}