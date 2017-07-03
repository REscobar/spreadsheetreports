namespace SpreadSheetsReports.ReportModel
{
    using System;
    using System.Runtime.Serialization;

    [DataContract]
    public abstract class PropertyBindingBase : IPropertyBinding
    {
        public string PropertyName { get; set; }

        public string Expression { get; set; }

        [IgnoreDataMember]
        public DataSourceBrowser DataSource { get; set; }

        protected internal abstract void PerformBind(ReportControl reportControl);

        public virtual PropertyBindingCollection Owner { get; set; }
    }
}
