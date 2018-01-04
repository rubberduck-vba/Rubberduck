using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public sealed class BinaryOpExpression : BoundExpression
    {
        public BinaryOpExpression(
            Declaration referencedDeclaration,
            ParserRuleContext context,
            IBoundExpression left,
            IBoundExpression right)
            : base(referencedDeclaration, ExpressionClassification.Value, context)
        {
            Left = left;
            Right = right;
        }

        public IBoundExpression Left { get; }

        public IBoundExpression Right { get; }
    }
}
