using System.Runtime.Serialization;

namespace SpreadSheetsReports.ReportModel
{
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