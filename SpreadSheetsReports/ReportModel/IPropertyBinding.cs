namespace SpreadSheetsReports.ReportModel
{
    public interface IPropertyBinding
    {
        string PropertyName { get; set; }

        string Expression { get; set; }

        DataSourceBrowser DataSource { get; set; }
    }
}