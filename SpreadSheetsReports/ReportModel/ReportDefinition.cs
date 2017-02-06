namespace SpreadSheetsReports.ReportModel
{
    using System.Collections.Generic;

    public class ReportDefinition : ReportControl
    {
        public List<Sheet> Sheets { get; set; }

        public object DataSource { get; set; }

        public void Render()
        {
            this.DoRender();
        }

        internal override void DoRender()
        {
            foreach (var sheet in this.Sheets)
            {
                if (sheet == null)
                {
                    continue;
                }

                sheet.DoRender();
            }
        }
    }
}
