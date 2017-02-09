namespace SpreadSheetsReports.Renderer
{
    using System.IO;
    using ReportModel;

    public interface IReportRenderer
    {
        Stream Render(ReportDefinition report);
    }
}
