namespace SpreadSheetsReports.ReportModel
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using SpreadSheetsReports.DocumentModel;
    using SpreadSheetsReports.ReportModel;

    public class Cell
    {
        public Cell()
        {
            this.Bindings = new CellBindingCollection();
        }

        public CellStyle Style { get; set; }

        public string ClassName { get; set; }

        public object Value { get; set; }

        public CellType Type { get; set; }

        public CellBindingCollection Bindings { get; private set; }

        public void Render()
        {
            foreach (var binding in this.Bindings)
            {
                binding.PerformBind(this);
            }
        }
    }
}
