using SpreadSheetsReports.DocumentModel;
using SpreadSheetsReports.ReportModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadSheetsReports
{
    public class Cell
    {
        public Cell()
        {
            this.Bindings = new CellBindingCollection();
        }
        public CellStyle Style { get; set; }
        public string ClassName { get; set; }
        public string Text { get; set; }
        public CellType Type { get; set; }

        public CellBindingCollection Bindings { get; private set; }
    }
}
