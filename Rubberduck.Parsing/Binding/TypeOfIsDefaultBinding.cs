using Antlr4.Runtime;

namespace Rubberduck.Parsing.Binding
{
    public sealed class TypeOfIsDefaultBinding : IExpressionBinding
    {
        private readonly ParserRuleContext _context;
        private readonly IExpressionBinding _expressionBinding;
        private readonly IExpressionBinding _typeExpressionBinding;

        public TypeOfIsDefaultBinding(
            ParserRuleContext context,
            IExpressionBinding expressionBinding,
            IExpressionBinding typeExpressionBinding)
        {
            _context = context;
            _expressionBinding = expressionBinding;
            _typeExpressionBinding = typeExpressionBinding;
        }

        public IBoundExpression Resolve()
        {
            var expr = _expressionBinding.Resolve();
            if (expr == null)
            {
                return null;
            }
            var typeExpr = _typeExpressionBinding.Resolve();
            if (typeExpr == null)
            {
                return null;
            }
            return new TypeOfIsExpression(null, _context, expr, typeExpr);
        }
    }
}
