using Antlr4.Runtime;

namespace Rubberduck.Parsing.Binding
{
    public sealed class UnaryOpDefaultBinding : IExpressionBinding
    {
        private readonly ParserRuleContext _context;
        private readonly IExpressionBinding _expr;

        public UnaryOpDefaultBinding(
            ParserRuleContext context,
            IExpressionBinding expr)
        {
            _context = context;
            _expr = expr;
        }

        public IBoundExpression Resolve()
        {
            // TODO: Allow broken trees?
            var expr = _expr.Resolve();
            if (expr == null)
            {
                return null;
            }
            return new UnaryOpExpression(expr.ReferencedDeclaration, _context, expr);
        }
    }
}
