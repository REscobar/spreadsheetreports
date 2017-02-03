namespace SpreadSheetsReports.ReportModel
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using SpreadSheetsReports.ReportModel;

    public class CellBindingCollection : List<CellBinding>, IList<CellBinding>
    {
        public void Add(string cellPropertyName, object dataSource, string dataMember)
        {
            if (string.IsNullOrWhiteSpace(cellPropertyName))
            {
                throw new ArgumentException(nameof(cellPropertyName));
            }

            if (dataSource == null)
            {
                throw new ArgumentNullException(nameof(dataSource));
            }

            if (string.IsNullOrEmpty(dataMember))
            {
                throw new ArgumentException(nameof(dataMember));
            }

            this.Add(new CellBinding(cellPropertyName, dataSource, dataMember));
        }

    }
}
