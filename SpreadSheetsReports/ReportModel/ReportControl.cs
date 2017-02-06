namespace SpreadSheetsReports.ReportModel
{
    public abstract class ReportControl
    {
        public ReportControl()
        {
            this.Bindings = new PropertyBindingCollection();
        }

        public PropertyBindingCollection Bindings { get; set; }

        internal virtual void DoRender()
        {
            foreach (var binding in this.Bindings)
            {
                binding.PerformBind(this);
            }
        }
    }
}