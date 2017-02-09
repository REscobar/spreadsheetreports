namespace SpreadSheetsReports.ReportModel
{
    using System;
    using DocumentModel;
    using Renderer;

    public class Cell : ReportControl
    {
        public Cell()
        {
            this.Style = new CellStyle();
        }

        public CellStyle Style { get; set; }

        public string ClassName { get; set; }

        public object Value { get; set; }

        public CellType Type { get; set; }

        protected override void DoRender(IReportRenderer renderer)
        {
            throw new NotImplementedException();
        }
    }
}
