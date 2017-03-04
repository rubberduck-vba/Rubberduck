using Antlr4.Runtime;

namespace Rubberduck.Parsing.Binding
{
    public sealed class BinaryOpDefaultBinding : IExpressionBinding
    {
        private readonly ParserRuleContext _context;
        private readonly IExpressionBinding _left;
        private readonly IExpressionBinding _right;

        public BinaryOpDefaultBinding(
            ParserRuleContext context,
            IExpressionBinding left,
            IExpressionBinding right)
        {
            _context = context;
            _left = left;
            _right = right;
        }

        public IBoundExpression Resolve()
        {
            var leftExpr = _left.Resolve();
            var rightExpr = _right.Resolve();
            if (leftExpr.Classification == ExpressionClassification.ResolutionFailed || rightExpr.Classification == ExpressionClassification.ResolutionFailed)
            {
                var failedExpr = new ResolutionFailedExpression();
                failedExpr.AddSuccessfullyResolvedExpression(leftExpr);
                failedExpr.AddSuccessfullyResolvedExpression(rightExpr);
                return failedExpr;
            }
            return new BinaryOpExpression(null, _context, leftExpr, rightExpr);
        }
    }
}
