namespace SpreadSheetsReports.ReportModel
{
    public interface ICellBinding
    {
        string CellPropertyName { get; set; }

        string DataMember { get; set; }

        object DataSource { get; set; }
    }
}