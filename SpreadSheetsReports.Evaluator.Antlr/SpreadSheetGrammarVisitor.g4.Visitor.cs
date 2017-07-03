namespace SpreadSheetsReports.Evaluator.Antlr
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Linq.Expressions;
    using System.Text;
    using System.Threading.Tasks;
    using Antlr4.Runtime.Misc;

    internal class SpreadSheetGrammarVisitor : SpreadSheetGrammarBaseVisitor<object>
    {
        private EvaluationContext context;
        private ParameterExpression @this;
        private ParameterExpression @param;

        public SpreadSheetGrammarVisitor(EvaluationContext context)
        {
            this.context = context;
            this.@this = Expression.Parameter(context.Target.GetType(), "this");
            this.@param = Expression.Parameter(context.Source.GetType(), "param");
        }

        public override object VisitCompilationUnit([NotNull] SpreadSheetGrammarParser.CompilationUnitContext context)
        {
            var parsed = base.VisitCompilationUnit(context);

            var block = Expression.Block(parsed as IEnumerable<Expression>);

            var lambda = Expression.Lambda(block, this.@this, this.@param);
            return lambda.Compile();
        }

        public override object VisitStatements([NotNull] SpreadSheetGrammarParser.StatementsContext context)
        {
            if (context == null)
            {
                return null;
            }

            var head = this.VisitStatements(context.statements());
            var tail = this.VisitStatement(context.statement());

            var list = head as List<Expression> ?? new List<Expression>();

            list.Add(tail as Expression);
            return list;
        }

        public override object VisitStatement([NotNull] SpreadSheetGrammarParser.StatementContext context)
        {
            var isBlock = context.RBRACE() != null;
            if (isBlock)
            {
                return Expression.Block(this.Visit(context.statements()) as IEnumerable<Expression>);
            }

            return this.Visit(context.GetChild(0));
        }

        public override object VisitAssignStatement([NotNull] SpreadSheetGrammarParser.AssignStatementContext context)
        {
            var memberAccess = this.VisitMemberAccessExpression(context.memberAccessExpression()) as Expression;
            var expression = this.VisitExpression(context.expression()) as Expression;

            var right = Expression.Convert(expression, memberAccess.Type);

            return Expression.Assign(memberAccess, right);
        }

        public override object VisitIfStatement([NotNull] SpreadSheetGrammarParser.IfStatementContext context)
        {
            var condition = this.Visit(context.expression()) as Expression;
            var trueStatment = this.Visit(context.statement(0)) as Expression;
            var falseStatementContext = context.statement(1);
            Expression falseStatement;
            var hasElse = falseStatementContext != null;
            if (hasElse)
            {
                falseStatement = this.Visit(falseStatementContext) as Expression;
                return Expression.IfThenElse(condition, trueStatment, falseStatement);
            }
            else
            {
                return Expression.IfThen(condition, trueStatment);
            }
        }

        public override object VisitStringLiteral([NotNull] SpreadSheetGrammarParser.StringLiteralContext context)
        {
            return Expression.Constant(context.GetText());
        }

        public override object VisitRelationalExpression([NotNull] SpreadSheetGrammarParser.RelationalExpressionContext context)
        {
            var leftContext = context.relationalExpression();

            Expression left = null;
            if (leftContext != null)
            {
                left = this.Visit(leftContext) as Expression;
            }

            var right = this.Visit(context.addExpression()) as Expression;

            if (left == null)
            {
                return right;
            }

            var @operator = context.GetChild(1).GetText();
            switch (@operator)
            {
                case "<":
                    return Expression.LessThan(left, right);
                case "<=":
                    return Expression.LessThanOrEqual(left, right);
                case ">":
                    return Expression.GreaterThan(left, right);
                case ">=":
                    return Expression.GreaterThanOrEqual(left, right);
                case "==":
                    return Expression.Equal(left, right);
                case "!=":
                    return Expression.NotEqual(left, right);
                default:
                    throw new InvalidOperationException();
            }
        }

        public override object VisitExpression([NotNull] SpreadSheetGrammarParser.ExpressionContext context)
        {
            var parsed = base.VisitExpression(context);
            return parsed;
        }

        public override object VisitLiteralExpression([NotNull] SpreadSheetGrammarParser.LiteralExpressionContext context)
        {
            var isGrouped = context.LPAREN() != null;
            if (isGrouped)
            {
                return this.Visit(context.expression());
            }

            var parsed = base.VisitLiteralExpression(context);
            return parsed;
        }

        public override object VisitMemberAccessExpression([NotNull] SpreadSheetGrammarParser.MemberAccessExpressionContext context)
        {
            if (context == null)
            {
                return null;
            }

            var parts = this.VisitMemberAccessExpression(context.memberAccessExpression());
            var identifier = this.VisitIdentifier(context.identifier());
            return this.GenerateMemberAccessExpression(identifier.ToString(), parts);
        }

        public override object VisitMultExpression([NotNull] SpreadSheetGrammarParser.MultExpressionContext context)
        {
            var leftContext = context.multExpression();

            Expression left = null;
            if (leftContext != null)
            {
                left = this.Visit(leftContext) as Expression;
            }

            var right = this.Visit(context.negateExpression()) as Expression;

            if (left == null)
            {
                return right;
            }

            var @operator = context.GetChild(1);
            if (@operator.GetText() == "*")
            {
                return Expression.Multiply(left, right);
            }
            else
            {
                return Expression.Divide(left, right);
            }
        }

        public override object VisitAddExpression([NotNull] SpreadSheetGrammarParser.AddExpressionContext context)
        {
            var leftContext = context.addExpression();

            Expression left = null;
            if (leftContext != null)
            {
                left = this.Visit(leftContext) as Expression;
            }

            var right = this.Visit(context.multExpression()) as Expression;

            if (left == null)
            {
                return right;
            }

            var @operator = context.GetChild(1);
            if (@operator.GetText() == "+")
            {
                return Expression.Add(left, right);
            }
            else
            {
                return Expression.Subtract(left, right);
            }
        }

        public override object VisitNegateExpression([NotNull] SpreadSheetGrammarParser.NegateExpressionContext context)
        {
            var parsed = this.Visit(context.primaryExpression()) as Expression;
            if (context.GetChild(0).GetText() == "-")
            {
                parsed = Expression.Negate(parsed);
            }

            return parsed;
        }

        protected override object AggregateResult(object aggregate, object nextResult)
        {
            return base.AggregateResult(aggregate, nextResult);
            if (aggregate == null)
            {
                return new[] { nextResult };
            }
            else
            {
                return ((object[])aggregate).Concat(new[] { nextResult }).ToArray();
            }
        }

        public override object VisitIdentifier([NotNull] SpreadSheetGrammarParser.IdentifierContext context)
        {
            var parsed = context.GetText();
            return parsed;
        }

        public override object VisitNumberLiteral([NotNull] SpreadSheetGrammarParser.NumberLiteralContext context)
        {
            return Expression.Constant(Convert.ToDecimal(context.GetText()));
        }

        private object GenerateMemberAccessExpression(string member, object target)
        {
            if (target == null)
            {
                if (member == "this")
                {
                    return this.@this;
                }
                else if (member == "param")
                {
                    return this.@param;
                }
                else
                {
                    return Type.GetType(member) ?? Type.GetType("System." + member);
                }
            }

            var targetType = target as Type;
            if (targetType != null)
            {
                return Expression.MakeMemberAccess(null, targetType.GetMember(member).First());
            }
            else
            {
                var newTargetExpression = target as Expression;
                return Expression.MakeMemberAccess(newTargetExpression, newTargetExpression.Type.GetMember(member).FirstOrDefault());
            }
        }
    }
}
