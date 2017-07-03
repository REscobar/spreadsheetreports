namespace SpreadSheetsReports.ReportModel
{
    using System.Collections.Generic;
    using System.Linq;
    using DocumentModel;

    public class Row : ReportControl, IRowGenerator
    {
        public List<ICellGenerator> Cells { get; set; }

        public float? Height { get; set; }

        public CellStyle Style { get; set; }

        public DocumentModel.Row Generate()
        {
            this.Databind();
            var row = new DocumentModel.Row();
            row.Cells = this.GetCells().ToList();
            row.Height = this.Height;
            row.Style = this.Style;
            return row;
        }

        private IEnumerable<DocumentModel.Cell> GetCells()
        {
            foreach (var cell in this.Cells)
            {
                yield return cell?.Generate();
            }
        }
    }
}
