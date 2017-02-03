namespace SpreadSheetsReports
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using ReportModel;
    using System.IO;
    public interface IReportRenderer
    {
        Stream Render(ReportDefinition report);
    }
}
