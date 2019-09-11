using Antlr4.Runtime;

namespace Rubberduck.Parsing.Binding
{
    public sealed class MissingArgumentExpression : BoundExpression
    {
        public MissingArgumentExpression(
            ExpressionClassification classification,
            ParserRuleContext context)
            : base(null, classification, context)
        {}
    }
}