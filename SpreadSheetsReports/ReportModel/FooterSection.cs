namespace SpreadSheetsReports.ReportModel
{
    public class FooterSection : ReportControl
    {
        public RowCollection Rows { get; set; }

        internal override void DoRender()
        {
            base.DoRender();

            foreach (var row in this.Rows)
            {
                if (row == null)
                {
                    continue;
                }

                row.DoRender();
            }
        }
    }
}