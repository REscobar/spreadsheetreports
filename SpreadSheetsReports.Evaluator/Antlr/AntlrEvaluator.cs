using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadSheetsReports.Evaluator.Antlr
{
    public class AntlrEvaluator : IEvaluator
    {
        public void Evaluate()
        {
            var test = new System.IO.StringReader("this.ts = 0");
            var parser = new SpreadSheetGrammarParser(new Antlr4.Runtime.UnbufferedTokenStream( new SpreadSheetGrammarLexer( new Antlr4.Runtime.UnbufferedCharStream(test))));

            //parser.statement

            throw new NotImplementedException();
        }
    }
}
