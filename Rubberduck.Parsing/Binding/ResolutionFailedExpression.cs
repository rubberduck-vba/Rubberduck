using System.Collections.Generic;

namespace Rubberduck.Parsing.Binding
{
    public sealed class ResolutionFailedExpression : BoundExpression
    {
        private readonly List<IBoundExpression> _successfullyResolvedExpressions;

        public ResolutionFailedExpression()
            : base(null, ExpressionClassification.ResolutionFailed, null)
        {
            _successfullyResolvedExpressions = new List<IBoundExpression>();
        }

        public IReadOnlyList<IBoundExpression> SuccessfullyResolvedExpressions => _successfullyResolvedExpressions;

        public void AddSuccessfullyResolvedExpression(IBoundExpression expression)
        {
            _successfullyResolvedExpressions.Add(expression);
        }
    }
}
