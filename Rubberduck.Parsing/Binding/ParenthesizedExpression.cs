using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public sealed class ParenthesizedExpression : BoundExpression
    {
        public ParenthesizedExpression(
            Declaration referencedDeclaration,
            ParserRuleContext context,
            IBoundExpression expression)
            : base(referencedDeclaration, ExpressionClassification.Value, context)
        {
            Expression = expression;
        }

        public IBoundExpression Expression { get; }
    }
}
