namespace SpreadSheetsReports.ReportModel
{
    using System;
    using DocumentModel;

    public class Cell : ReportControl, ICellGenerator
    {
        public Cell()
        {
            this.Style = new CellStyle();
        }

        public CellStyle Style { get; set; }

        public string ClassName { get; set; }

        public object Value { get; set; }

        public CellType Type { get; set; }

        public DocumentModel.Cell Generate()
        {
            this.Databind();
            var cell = new DocumentModel.Cell();

            cell.ClassName = this.ClassName;
            cell.Value = this.Value;
            cell.Type = this.Type;
            cell.Style = this.Style;

            return cell;
        }
    }
}
