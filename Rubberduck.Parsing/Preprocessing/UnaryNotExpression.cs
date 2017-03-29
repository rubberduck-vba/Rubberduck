using System;

namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class UnaryNotExpression : Expression
    {
        private readonly IExpression _expression;

        public UnaryNotExpression(IExpression expression)
        {
            _expression = expression;
        }

        public override IValue Evaluate()
        {
            var operand = _expression.Evaluate();
            if (operand == null)
            {
                return null;
            }
            else if (operand.ValueType == ValueType.Bool)
            {
                return new BoolValue(!operand.AsBool);
            }
            else
            {
                var coerced = operand.AsDecimal;
                return new DecimalValue(~Convert.ToInt64(coerced));
            }
        }
    }
}
