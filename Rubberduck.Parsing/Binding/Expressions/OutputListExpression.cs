using System.Collections.Generic;
using Antlr4.Runtime;

namespace Rubberduck.Parsing.Binding
{
    public sealed class OutputListExpression : BoundExpression
    {
        public OutputListExpression(
            ParserRuleContext context,
            IReadOnlyCollection<IBoundExpression> itemExpressions)
            : base(null, ExpressionClassification.Value, context)
        {
            ItemExpressions = itemExpressions;
        }

        public IReadOnlyCollection<IBoundExpression> ItemExpressions { get; }
    }
}