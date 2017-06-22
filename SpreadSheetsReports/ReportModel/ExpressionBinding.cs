namespace SpreadSheetsReports.ReportModel
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;

    public class ExpressionBinding : PropertyBindingBase
    {
        public string PropertyName { get; set; }

        public string Expression { get; set; }

        public DataSourceBrowser DataSource { get; set; }

        public PropertyBindingCollection Owner { get; set; }

        protected internal override void PerformBind(ReportControl reportControl)
        {
            throw new NotImplementedException();
        }
    }
}
