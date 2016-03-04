namespace Rubberduck.Parsing.Preprocessing
{
    public sealed class LogicalLessThanExpression : Expression
    {
        private readonly IExpression _left;
        private readonly IExpression _right;

        public LogicalLessThanExpression(IExpression left, IExpression right)
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
            if (left.ValueType == ValueType.String && right.ValueType == ValueType.String)
            {
                var leftValue = left.AsString;
                var rightValue = right.AsString;
                return new BoolValue(string.CompareOrdinal(leftValue, rightValue) < 0);
            }
            else if (left.ValueType == ValueType.String && right.ValueType == ValueType.Empty)
            {
                return new BoolValue(string.CompareOrdinal(left.AsString, right.AsString) < 0);
            }
            else if (right.ValueType == ValueType.String && left.ValueType == ValueType.Empty)
            {
                return new BoolValue(string.CompareOrdinal(right.AsString, left.AsString) < 0);
            }
            else
            {
                var leftValue = left.AsDecimal;
                var rightValue = right.AsDecimal;
                return new BoolValue(leftValue < rightValue);
            }
        }
    }
}
