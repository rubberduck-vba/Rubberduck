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
            var expr = _expr.Resolve();
            return expr.Classification == ExpressionClassification.ResolutionFailed 
                ? expr 
                : new UnaryOpExpression(expr.ReferencedDeclaration, _context, expr);
        }
    }
}
