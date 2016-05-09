using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public sealed class ParenthesizedDefaultBinding : IExpressionBinding
    {
        private readonly ParserRuleContext _context;
        private readonly IExpressionBinding _expressionBinding;

        public ParenthesizedDefaultBinding(
            ParserRuleContext context,
            IExpressionBinding expressionBinding)
        {
            _context = context;
            _expressionBinding = expressionBinding;
        }

        public IBoundExpression Resolve()
        {
            var expr = _expressionBinding.Resolve();
            if (expr == null)
            {
                return null;
            }
            return new ParenthesizedExpression(expr.ReferencedDeclaration, _context, expr);
        }
    }
}
