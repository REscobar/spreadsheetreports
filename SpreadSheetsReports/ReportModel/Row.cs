namespace SpreadSheetsReports.ReportModel
{
    using System.Collections.Generic;
    using DocumentModel;

    public class Row : ReportControl, IRowGenerator
    {
        public List<Cell> Cells { get; set; }

        public float? Height { get; set; }

        public CellStyle Style { get; set; }

        public DocumentModel.Row Generate()
        {
            var row = new DocumentModel.Row();
            row.Cells = this.GetCells();
            row.Height = this.Height;
            row.Style = this.Style;
            return row;
        }

        protected override void DoRender()
        {
            foreach (var cell in this.Cells)
            {
                if (cell == null)
                {
                    continue;
                }

                cell.Render();
            }
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
