using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadSheetsReports
{
    public class ReportSection 
    {
        public HeaderSection Header { get; set; }
        public ReportSection Detail { get; set; }
        public FooterSection Footer { get; set; }
    }
}
