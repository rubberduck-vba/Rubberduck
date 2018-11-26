using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public sealed class BinaryOpExpression : BoundExpression
    {
        private readonly IBoundExpression _left;
        private readonly IBoundExpression _right;

        public BinaryOpExpression(
            Declaration referencedDeclaration,
            ParserRuleContext context,
            IBoundExpression left,
            IBoundExpression right)
            : base(referencedDeclaration, ExpressionClassification.Value, context)
        {
            _left = left;
            _right = right;
        }

        public IBoundExpression Left
        {
            get
            {
                return _left;
            }
        }

        public IBoundExpression Right
        {
            get
            {
                return _right;
            }
        }
    }
}
