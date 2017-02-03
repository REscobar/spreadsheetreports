namespace SpreadSheetsReports.ReportModel
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;

    public class ReportSection
    {
        public HeaderSection Header { get; set; }

        public ReportSection SubSection { get; set; }

        public FooterSection Footer { get; set; }
    }
}
