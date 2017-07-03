using System.Runtime.Serialization;

namespace SpreadSheetsReports.ReportModel
{
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