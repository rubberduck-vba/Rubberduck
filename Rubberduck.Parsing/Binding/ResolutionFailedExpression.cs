using System.Collections.Generic;

namespace Rubberduck.Parsing.Binding
{
    public sealed class ResolutionFailedExpression : BoundExpression
    {
        private readonly List<IBoundExpression> _successfullyResolvedExpressions = new List<IBoundExpression>();

        public ResolutionFailedExpression()
            : base(null, ExpressionClassification.ResolutionFailed, null)
        {
            _successfullyResolvedExpressions = new List<IBoundExpression>();
        }

        public IReadOnlyList<IBoundExpression> SuccessfullyResolvedExpressions
        {
            get
            {
                return _successfullyResolvedExpressions;
            }
        }

        public void AddSuccessfullyResolvedExpression(IBoundExpression expression)
        {
            _successfullyResolvedExpressions.Add(expression);
        }
    }
}
