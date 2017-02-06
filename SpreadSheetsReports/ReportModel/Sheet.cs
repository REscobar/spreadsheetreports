namespace SpreadSheetsReports.ReportModel
{
    public class Sheet : ReportControl
    {
        public ReportSection Content { get; set; }

        public string Name { get; set; }

        internal override void DoRender()
        {
            base.DoRender();

            if (this.Content == null)
            {
                return;
            }

            this.Content.DoRender();
        }
    }
}
