namespace SpreadSheetsReports.ReportModel
{
    public abstract class ReportControl
    {
        public ReportControl()
        {
            this.Bindings = new PropertyBindingCollection();
        }

        public PropertyBindingCollection Bindings { get; set; }

        public string DataMember { get; set; }

        public DataSourceBrowser DataSource { get; set; }

        //public virtual void Render()
        //{
        //    this.Databind();
        //    this.DoRender();
        //}

        //protected abstract void DoRender();

        protected virtual void Databind()
        {
            foreach (var binding in this.Bindings)
            {
                binding.PerformBind(this);
            }
        }
    }
}