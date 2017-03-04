using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public sealed class ParenthesizedExpression : BoundExpression
    {
        private readonly IBoundExpression _expression;

        public ParenthesizedExpression(
            Declaration referencedDeclaration,
            ParserRuleContext context,
            IBoundExpression expression)
            : base(referencedDeclaration, ExpressionClassification.Value, context)
        {
            _expression = expression;
        }

        public IBoundExpression Expression
        {
            get
            {
                return _expression;
            }
        }
    }
}
