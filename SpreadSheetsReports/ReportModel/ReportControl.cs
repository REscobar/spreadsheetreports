namespace SpreadSheetsReports.ReportModel
{
    using System.Runtime.Serialization;

    [DataContract]
    public abstract class ReportControl
    {
        public ReportControl()
        {
            this.Bindings = new PropertyBindingCollection();
        }

        [DataMember]
        public PropertyBindingCollection Bindings { get; set; }

        protected virtual void Databind()
        {
            foreach (var binding in this.Bindings)
            {
                binding.PerformBind(this);
            }
        }
    }
}