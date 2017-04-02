using System;

namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class LogicalImpExpression : Expression
    {
        private readonly IExpression _left;
        private readonly IExpression _right;

        public LogicalImpExpression(IExpression left, IExpression right)
        {
            _left = left;
            _right = right;
        }

        public override IValue Evaluate()
        {
            var left = _left.Evaluate();
            var right = _right.Evaluate();
            if (left == null && right == null)
            {
                return null;
            }
            else if (left != null && left.ValueType == ValueType.Bool && right != null && right.ValueType == ValueType.Bool)
            {
                var leftNumber = Convert.ToInt64(left.AsDecimal);
                var rightNumber = Convert.ToInt64(right.AsDecimal);
                var result = (decimal)(~leftNumber | rightNumber);
                return new BoolValue(new DecimalValue(result).AsBool);
            }
            else if (left == null && right.AsDecimal== 0)
            {
                return null;
            }
            else if (left == null)
            {
                return right;
            }
            else if (left.AsDecimal== -1)
            {
                return null;
            }
            else if (right == null)
            {
                return new DecimalValue(~Convert.ToInt64(left.AsDecimal) | 0);
            }
            else
            {
                var leftNumber = Convert.ToInt64(left.AsDecimal);
                var rightNumber = Convert.ToInt64(right.AsDecimal);
                return new DecimalValue(~leftNumber | rightNumber);
            }
        }
    }
}
