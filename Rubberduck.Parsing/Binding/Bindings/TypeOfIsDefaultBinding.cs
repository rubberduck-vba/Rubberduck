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
            var typeExpr = _typeExpressionBinding.Resolve();
            if (expr.Classification == ExpressionClassification.ResolutionFailed || typeExpr.Classification == ExpressionClassification.ResolutionFailed)
            {
                var failedExpr = new ResolutionFailedExpression();
                failedExpr.AddSuccessfullyResolvedExpression(expr);
                failedExpr.AddSuccessfullyResolvedExpression(typeExpr);
                return failedExpr;
            }
            return new TypeOfIsExpression(null, _context, expr, typeExpr);
        }
    }
}
