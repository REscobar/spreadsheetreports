namespace SpreadSheetsReports.ReportModel
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;

    public class ReportDefinition
    {
        public IEnumerable<Sheet> Sheets { get; set; }

        public object DataSource { get; set; }
    }
}
