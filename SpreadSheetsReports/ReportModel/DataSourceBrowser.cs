using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadSheetsReports.ReportModel
{
    public abstract class DataSourceBrowser : IDataSourceBrowser
    {
        public abstract object Current { get; }

        public abstract object GetValue(string property);

        public abstract bool MoveNext();
    }
}
