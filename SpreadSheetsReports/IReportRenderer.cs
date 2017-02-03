﻿namespace SpreadSheetsReports
{
    using ReportModel;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;

    public interface IReportRenderer
    {
        void Render(ReportDefinition report);
    }
}
