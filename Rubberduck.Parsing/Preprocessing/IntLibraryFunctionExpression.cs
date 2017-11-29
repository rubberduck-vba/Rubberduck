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
            switch (expr.ValueType)
            {
                case ValueType.Decimal:
                    return new DecimalValue(Math.Truncate(expr.AsDecimal));
                case ValueType.String:
                    return new DecimalValue(Math.Truncate(expr.AsDecimal));
                case ValueType.Date:
                    var truncated = new DecimalValue(Math.Truncate(expr.AsDecimal));
                    return new DateValue(truncated.AsDate);
                default:
                    return new DecimalValue(expr.AsDecimal);
            }
        }
    }
}
