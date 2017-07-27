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
        [IgnoreDataMember]
        public IDataSourceBrowser DataSource { get; set; }

        public PropertyBindingCollection Owner { get; set; }

        [DataMember]
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
