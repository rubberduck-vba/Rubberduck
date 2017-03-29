using System;

namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class LogicalAndExpression : Expression
    {
        private readonly IExpression _left;
        private readonly IExpression _right;

        public LogicalAndExpression(IExpression left, IExpression right)
        {
            _left = left;
            _right = right;
        }

        public override IValue Evaluate()
        {
            var left = _left.Evaluate();
            var right = _right.Evaluate();
            if (left != null && left.AsDecimal== 0 && right == null)
            {
                return new DecimalValue(0);
            }
            else if (right != null && right.AsDecimal== 0 && left == null)
            {
                return new DecimalValue(0);
            }
            else if (left == null || right == null)
            {
                return null;
            }
            else if (left.ValueType == ValueType.Bool && right.ValueType == ValueType.Bool)
            {
                var leftNumber = Convert.ToInt64(left.AsDecimal);
                var rightNumber = Convert.ToInt64(right.AsDecimal);
                return new BoolValue(new DecimalValue(leftNumber & rightNumber).AsBool);
            }
            else
            {
                var leftNumber = Convert.ToInt64(left.AsDecimal);
                var rightNumber = Convert.ToInt64(right.AsDecimal);
                return new DecimalValue(leftNumber & rightNumber);
            }
        }
    }
}
