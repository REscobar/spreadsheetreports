namespace SpreadSheetsReports.ReportModel
{
    public class ReportSection : ReportControl
    {
        public HeaderSection Header { get; set; }

        public ReportSection SubSection { get; set; }

        public FooterSection Footer { get; set; }

        internal override void DoRender()
        {
            base.DoRender();
            if (this.Header != null)
            {
                this.Header.DoRender();
            }

            if (this.SubSection != null)
            {
                this.SubSection.DoRender();
            }

            if (this.SubSection != null)
            {
                this.Footer.DoRender();
            }
        }
    }
}
