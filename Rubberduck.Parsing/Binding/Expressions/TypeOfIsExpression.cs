using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public sealed class TypeOfIsExpression : BoundExpression
    {
        private readonly IBoundExpression _expression;
        private readonly IBoundExpression _typeExpression;

        public TypeOfIsExpression(
            Declaration referencedDeclaration,
            ParserRuleContext context,
            IBoundExpression expression,
            IBoundExpression typeExpression)
            : base(referencedDeclaration, ExpressionClassification.Value, context)
        {
            _expression = expression;
            _typeExpression = typeExpression;
        }

        public IBoundExpression Expression
        {
            get
            {
                return _expression;
            }
        }

        public IBoundExpression TypeExpression
        {
            get
            {
                return _typeExpression;
            }
        }
    }
}
