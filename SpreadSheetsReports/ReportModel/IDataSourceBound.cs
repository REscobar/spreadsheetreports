namespace SpreadSheetsReports.ReportModel
{
    public interface IDataSourceBound
    {
        string DataMember { get; set; }

        DataSourceBrowser DataSource { get; set; }
    }
}