using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadSheetsReports.ReportModel
{
    public class CellBinding
    {
        public string CellPropertyName { get; set; }
        public string DataMember { get; set; }
        public object DataSource { get; set; }

        public CellBinding(string cellPorpertyName, object dataSource, string dataMember)
        {
            this.CellPropertyName = cellPorpertyName;
            this.DataSource = dataSource;
            this.DataMember = dataMember;
        }
    }
}
