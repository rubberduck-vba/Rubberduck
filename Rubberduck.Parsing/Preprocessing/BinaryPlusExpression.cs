namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class BinaryPlusExpression : Expression
    {
        private readonly IExpression _left;
        private readonly IExpression _right;

        public BinaryPlusExpression(IExpression left, IExpression right)
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
            if (left.ValueType == ValueType.String || right.ValueType == ValueType.String)
            {
                return new StringValue(left.AsString+ right.AsString);
            }
            else if (left.ValueType == ValueType.Date || right.ValueType == ValueType.Date)
            {
                decimal leftValue = left.AsDecimal;
                decimal rightValue = right.AsDecimal;
                decimal sum = leftValue + rightValue;
                try
                {
                    return new DateValue(new DecimalValue(sum).AsDate);
                }
                catch
                {
                    return new DecimalValue(sum);
                }
            }
            else
            {
                var leftNumber = left.AsDecimal;
                var rightNumber = right.AsDecimal;
                return new DecimalValue(leftNumber + rightNumber);
            }
        }
    }
}
