using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadSheetsReports
{
    public interface IReportRenderer
    {
        void Render(ReportDefinition report);
    }
}
