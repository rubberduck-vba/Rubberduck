using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public sealed class UnaryOpExpression : BoundExpression
    {
        private readonly IBoundExpression _expr;

        public UnaryOpExpression(
            Declaration referencedDeclaration,
            ParserRuleContext context,
            IBoundExpression expr)
            : base(referencedDeclaration, ExpressionClassification.Value, context)
        {
            _expr = expr;
        }

        public IBoundExpression Expr
        {
            get
            {
                return _expr;
            }
        }
    }
}
