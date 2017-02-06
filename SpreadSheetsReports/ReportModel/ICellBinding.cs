namespace SpreadSheetsReports.ReportModel
{
    public interface IPropertyBinding
    {
        string PropertyName { get; set; }

        string DataMember { get; set; }

        object DataSource { get; set; }
    }
}