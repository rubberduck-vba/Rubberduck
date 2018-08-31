using System;

namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class IntLibraryFunctionExpression : Expression
    {
        private readonly IExpression _expression;

        public IntLibraryFunctionExpression(IExpression expression)
        {
            _expression = expression;
        }

        public override IValue Evaluate()
        {
            var expr = _expression.Evaluate();
            if (expr == null)
            {
                return null;
            }
            if (expr.ValueType == ValueType.Decimal)
            {
                return new DecimalValue(Math.Truncate(expr.AsDecimal));
            }
            else if (expr.ValueType == ValueType.String)
            {
                return new DecimalValue(Math.Truncate(expr.AsDecimal));
            }
            else if (expr.ValueType == ValueType.Date)
            {
                var truncated = new DecimalValue(Math.Truncate(expr.AsDecimal));
                return new DateValue(truncated.AsDate);
            }
            else
            {
                return new DecimalValue(expr.AsDecimal);
            }
        }
    }
}
