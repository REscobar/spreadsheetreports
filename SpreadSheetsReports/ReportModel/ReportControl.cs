using SpreadSheetsReports.Renderer;

namespace SpreadSheetsReports.ReportModel
{
    public abstract class ReportControl
    {
        public ReportControl()
        {
            this.Bindings = new PropertyBindingCollection();
        }

        public PropertyBindingCollection Bindings { get; set; }

        public virtual void Render(IReportRenderer renderer)
        {
            this.Databind();
            this.DoRender(renderer);
        }

        protected abstract void DoRender(IReportRenderer renderer);

        internal virtual void Databind()
        {
            foreach (var binding in this.Bindings)
            {
                binding.PerformBind(this);
            }
        }
    }
}