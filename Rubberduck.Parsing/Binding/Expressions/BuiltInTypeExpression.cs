using Antlr4.Runtime;

namespace Rubberduck.Parsing.Binding
{
    public sealed class BuiltInTypeExpression : BoundExpression
    {
        public BuiltInTypeExpression(ParserRuleContext context)
            : base(null, ExpressionClassification.Type, context)
        {
        }
    }
}