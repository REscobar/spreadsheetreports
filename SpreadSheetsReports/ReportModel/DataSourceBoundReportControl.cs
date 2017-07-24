namespace SpreadSheetsReports.ReportModel
{
    using System.Runtime.Serialization;

    [DataContract]
    public abstract class DataSourceBoundReportControl : ReportControl, IDataSourceBound
    {
        public DataSourceBoundReportControl()
            : base()
        {
        }

        [DataMember]
        public string DataMember { get; set; }

        [IgnoreDataMember]
        public DataSourceBrowser DataSource { get; set; }
    }
}