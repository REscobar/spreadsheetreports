namespace SpreadSheetsReports.ReportModel
{
    public abstract class ReportControl
    {
        public ReportControl()
        {
            this.Bindings = new PropertyBindingCollection();
        }

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