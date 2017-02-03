namespace SpreadSheetsReports.ReportModel
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;

    public class CellBinding : ICellBinding
    {
        public CellBinding()
        {

        }

        public CellBinding(string cellPorpertyName, object dataSource, string dataMember)
        {
            this.CellPropertyName = cellPorpertyName;
            this.DataSource = dataSource;
            this.DataMember = dataMember;
        }

        public string CellPropertyName { get; set; }

        public string DataMember { get; set; }

        public object DataSource { get; set; }

        internal void PerformBind(Cell cell)
        {
            if (cell == null)
            {
                throw new ArgumentNullException(nameof(cell));
            }

            object value = this.ResolveValue();

            cell.GetType().GetProperty(this.CellPropertyName).SetMethod.Invoke(cell, new[] { value });
        }

        private object ResolveValue()
        {
            return this.DataSource.GetType().GetProperty(this.DataMember).GetValue(this.DataSource);
        }
    }
}
