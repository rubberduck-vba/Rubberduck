namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class BinaryMinusExpression : Expression
    {
        private readonly IExpression _left;
        private readonly IExpression _right;

        public BinaryMinusExpression(IExpression left, IExpression right)
        {
            _left = left;
            _right = right;
        }

        public override IValue Evaluate()
        {
            var left = _left.Evaluate();
            var right = _right.Evaluate();
            if (left == null || right == null)
            {
                return null;
            }
            else if (left.ValueType == ValueType.Date && right.ValueType == ValueType.Date)
            {
                // 5.6.9.3.3 - Effective value type exception.
                // If left + right are both Date then effective value type is double.
                decimal leftValue = left.AsDecimal;
                decimal rightValue = right.AsDecimal;
                decimal difference = leftValue - rightValue;
                return new DecimalValue(difference);
            }
            else if (left.ValueType == ValueType.Date || right.ValueType == ValueType.Date)
            {
                decimal leftValue = left.AsDecimal;
                decimal rightValue = right.AsDecimal;
                decimal difference = leftValue - rightValue;
                try
                {
                    return new DateValue(new DecimalValue(difference).AsDate);
                }
                catch
                {
                    return new DecimalValue(difference);
                }
            }
            else
            {
                decimal leftValue = left.AsDecimal;
                decimal rightValue = right.AsDecimal;
                return new DecimalValue(leftValue - rightValue);
            }
        }
    }
}
