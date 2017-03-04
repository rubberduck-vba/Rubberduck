using Antlr4.Runtime;

namespace Rubberduck.Parsing.Binding
{
    public sealed class LiteralExpression : BoundExpression
    {
        public LiteralExpression(ParserRuleContext context)
            : base(null, ExpressionClassification.Value, context)
        {
        }
    }
}
