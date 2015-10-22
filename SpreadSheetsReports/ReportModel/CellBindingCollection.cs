using SpreadSheetsReports.ReportModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;

namespace SpreadSheetsReports.ReportModel
{
    public class CellBindingCollection : List<CellBinding>
    {
        public void Add(string cellPropertName, object dataSource, string dataMember)
        {
            this.Add(new CellBinding(cellPropertName, dataSource, dataMember));
        }

    }
}
