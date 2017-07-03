namespace SpreadSheetsReports.Evaluator.Antlr
{
    using System;
    using Antlr4.Runtime;
    using System.IO;

    public class AntlrEvaluator : IEvaluator
    {
        public void Evaluate(EvaluationContext context)
        {
            using (var expressionReader = new StringReader(context.Expression))
            {
                var parser = new SpreadSheetGrammarParser(new UnbufferedTokenStream(new SpreadSheetGrammarLexer(new AntlrInputStream(expressionReader))));

                var compiled = parser.compilationUnit();
                var visitor = new SpreadSheetGrammarVisitor(context);
                var result = visitor.Visit(compiled) as Delegate;
                result.DynamicInvoke(context.Target, context.Source);
            }
        }
    }
}
