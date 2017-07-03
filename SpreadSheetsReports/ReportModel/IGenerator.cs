namespace SpreadSheetsReports.ReportModel
{
    public interface IGenerator<T>
    {
        T Generate();
    }
}