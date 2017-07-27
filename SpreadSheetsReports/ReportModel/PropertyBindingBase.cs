namespace SpreadSheetsReports.ReportModel
{
    using System;
    using System.Runtime.Serialization;

    [DataContract]
    public abstract class PropertyBindingBase : IPropertyBinding
    {
        [DataMember]
        public virtual string PropertyName { get; set; }

        [DataMember]
        public virtual string Expression { get; set; }

        [IgnoreDataMember]
        public DataSourceBrowser DataSource { get; set; }

        protected internal abstract void PerformBind(ReportControl reportControl);

        public virtual PropertyBindingCollection Owner { get; set; }
    }
}
