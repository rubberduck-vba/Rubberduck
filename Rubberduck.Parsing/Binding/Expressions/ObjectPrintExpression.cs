using Antlr4.Runtime;

namespace Rubberduck.Parsing.Binding
{
    public sealed class ObjectPrintExpression : BoundExpression
    {
        public ObjectPrintExpression(
            ParserRuleContext context,
            IBoundExpression printMethodExpression,
            IBoundExpression outputListBoundExpression)
            : base(null, ExpressionClassification.Subroutine, context)
        {
            PrintMethodExpressions = printMethodExpression;
            OutputListExpression = outputListBoundExpression;
        }

        public IBoundExpression PrintMethodExpressions { get; }
        public IBoundExpression OutputListExpression { get; }
    }
}