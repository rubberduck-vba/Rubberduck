using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;

namespace Rubberduck.Parsing.Binding
{
    public sealed class ResolutionFailedExpression : BoundExpression
    {
        private readonly List<IBoundExpression> _successfullyResolvedExpressions = new List<IBoundExpression>();

        public ResolutionFailedExpression(ParserRuleContext context, bool isDefaultMemberResolution = false)
            : base(null, ExpressionClassification.ResolutionFailed, context)
        {
            IsDefaultMemberResolution = isDefaultMemberResolution;
            IsJoinedExpression = false;
        }

        public ResolutionFailedExpression(ParserRuleContext context, IEnumerable<IBoundExpression> expressions)
            : base(null, ExpressionClassification.ResolutionFailed, context)
        {
            IsDefaultMemberResolution = false;
            IsJoinedExpression = true;

            AddSuccessfullyResolvedExpressions(expressions);
        }

        public IReadOnlyList<IBoundExpression> SuccessfullyResolvedExpressions => _successfullyResolvedExpressions;
        public bool IsDefaultMemberResolution { get; }
        public bool IsJoinedExpression { get; }

        public void AddSuccessfullyResolvedExpression(IBoundExpression expression)
        {
            _successfullyResolvedExpressions.Add(expression);
        }

        public void AddSuccessfullyResolvedExpressions(IEnumerable<IBoundExpression> expressions)
        {
            _successfullyResolvedExpressions.AddRange(expressions);
        }
    }

    public static class FailedResolutionExpressionExtensions
    {
        public static ResolutionFailedExpression JoinAsFailedResolution(this IBoundExpression expression, ParserRuleContext context, params IBoundExpression[] otherExpressions)
        {
            return expression.JoinAsFailedResolution(context, (IEnumerable<IBoundExpression>)otherExpressions);
        }

        public static ResolutionFailedExpression JoinAsFailedResolution(this IBoundExpression expression, ParserRuleContext context, IEnumerable<IBoundExpression> otherExpressions)
        {
            var expressions = otherExpressions.Concat(new[] {expression});
            return new ResolutionFailedExpression(context, expressions);
        }
    }
}
