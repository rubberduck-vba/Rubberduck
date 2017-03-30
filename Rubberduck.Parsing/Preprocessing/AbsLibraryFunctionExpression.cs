using System;

namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class AbsLibraryFunctionExpression : Expression
    {
        private readonly IExpression _expression;

        public AbsLibraryFunctionExpression(IExpression expression)
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
            if (expr.ValueType == ValueType.Date)
            {
                decimal exprValue = expr.AsDecimal;
                exprValue = Math.Abs(exprValue);
                try
                {
                    return new DateValue(new DecimalValue(exprValue).AsDate);
                }
                catch
                {
                    return new DecimalValue(exprValue);
                }
            }
            else
            {
                return new DecimalValue(Math.Abs(expr.AsDecimal));
            }
        }
    }
}
