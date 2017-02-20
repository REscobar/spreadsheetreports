namespace SpreadSheetsReports.ReportModel
{
    public interface IDataSourceBrowser
    {
        object Current { get; }

        object GetValue(string property);

        bool MoveNext();
    }
}