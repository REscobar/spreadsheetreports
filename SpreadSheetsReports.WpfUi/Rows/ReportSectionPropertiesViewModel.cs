using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadSheetsReports.WpfUi.Rows
{
    public class ReportSectionPropertiesViewModel
    {
        private readonly ReportSectionBinder binder;

        public ReportSectionPropertiesViewModel(ReportSectionBinder binder)
        {
            this.binder = binder;
        }

        public string DataMember { get; set; }

        public void Copy()
        {

        }
    }
}
