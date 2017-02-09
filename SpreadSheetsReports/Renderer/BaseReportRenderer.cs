namespace SpreadSheetsReports.Renderer
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using SpreadSheetsReports.ReportModel;

    public class BaseReportRenderer : IReportRenderer
    {
        public virtual Stream Render(ReportDefinition report)
        {
            throw new NotImplementedException();
        }
    }
}
