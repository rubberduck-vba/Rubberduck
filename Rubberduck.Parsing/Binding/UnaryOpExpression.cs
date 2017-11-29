using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public sealed class UnaryOpExpression : BoundExpression
    {
        public UnaryOpExpression(
            Declaration referencedDeclaration,
            ParserRuleContext context,
            IBoundExpression expr)
            : base(referencedDeclaration, ExpressionClassification.Value, context)
        {
            Expr = expr;
        }

        public IBoundExpression Expr { get; }
    }
}
