namespace SpreadSheetsReports.ReportModel
{
    internal interface IGenerator<T>
    {
        T Generate();
    }
}