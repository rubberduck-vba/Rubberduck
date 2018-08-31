namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class UnaryMinusExpression : Expression
    {
        private readonly IExpression _expression;

        public UnaryMinusExpression(IExpression expression)
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
            else if (operand.ValueType == ValueType.Date)
            {
                var value = operand.AsDecimal;
                value = -value;
                try
                {
                    return new DateValue(new DecimalValue(value).AsDate);
                }
                catch
                {
                    // 5.6.9.3.1: If overflow occurs during the coercion to Date, and the operand has a 
                    // declared type of Variant, the result is the Double value.
                    // We don't care about it being a Variant because if it's not a Variant it won't compile/run.
                    // We catch everything because the only case where the code is valid is that if this is an overflow.
                    return new DecimalValue(value);
                }
            }
            else
            {
                var value = operand.AsDecimal;
                return new DecimalValue(-value);
            }
        }
    }
}
