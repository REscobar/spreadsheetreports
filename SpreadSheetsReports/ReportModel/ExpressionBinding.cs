namespace SpreadSheetsReports.ReportModel
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Runtime.Serialization;
    using System.Text;
    using System.Threading.Tasks;
    using Evaluator;

    [DataContract]
    public class ExpressionBinding : PropertyBindingBase
    {
        public string PropertyName { get; set; }

        public string Expression { get; set; }

        [IgnoreDataMember]
        public IDataSourceBrowser DataSource { get; set; }

        public PropertyBindingCollection Owner { get; set; }

        public IEvaluator ExpressionEvaluator { get; set; }

        protected internal override void PerformBind(ReportControl reportControl)
        {
            this.ExpressionEvaluator.Evaluate(new EvaluationContext
            {
                Target = reportControl,
                Expression = this.Expression,
                Source = this.DataSource.Current
            });
        }
    }
}
