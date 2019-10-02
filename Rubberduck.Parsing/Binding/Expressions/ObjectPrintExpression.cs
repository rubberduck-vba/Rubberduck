using Antlr4.Runtime;

namespace Rubberduck.Parsing.Binding
{
    public sealed class ObjectPrintExpression : BoundExpression
    {
        public ObjectPrintExpression(
            ParserRuleContext context,
            IBoundExpression memberAccessExpression,
            IBoundExpression outputListBoundExpression)
            : base(null, ExpressionClassification.Subroutine, context)
        {
            MemberAccessExpressions = memberAccessExpression;
            OutputListExpression = outputListBoundExpression;
        }

        public IBoundExpression MemberAccessExpressions { get; }
        public IBoundExpression OutputListExpression { get; }
    }
}