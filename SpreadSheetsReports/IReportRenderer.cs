namespace SpreadSheetsReports
{
    using ReportModel;
    using System.IO;
    public interface IReportRenderer
    {
        Stream Render(ReportDefinition report);
    }
}
