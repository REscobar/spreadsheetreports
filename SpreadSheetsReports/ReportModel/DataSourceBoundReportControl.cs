namespace SpreadSheetsReports.ReportModel
{
    public abstract class DataSourceBoundReportControl : ReportControl, IDataSourceBound
    {
        public DataSourceBoundReportControl()
            : base()
        {
        }

        public string DataMember { get; set; }

        public DataSourceBrowser DataSource { get; set; }
    }
}