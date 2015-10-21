using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadSheetsReports
{
    public class ReportDefinition
    {
        public ReportSection Content { get; set; }

        public object DataSource { get; set; }
    }
}
