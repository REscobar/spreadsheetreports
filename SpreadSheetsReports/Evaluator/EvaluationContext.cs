namespace SpreadSheetsReports.Evaluator
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;

    public class EvaluationContext
    {
        public string Expression { get; set; }

        public object Source { get; set; }

        public object Target { get; set; }
    }
}
