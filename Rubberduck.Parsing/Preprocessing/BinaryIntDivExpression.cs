using System;

namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class BinaryIntDivExpression : Expression
    {
        private readonly IExpression _left;
        private readonly IExpression _right;

        public BinaryIntDivExpression(IExpression left, IExpression right)
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
            var leftValue = Convert.ToInt64(left.AsDecimal);
            var rightValue = Convert.ToInt64(right.AsDecimal);
            return new DecimalValue(Math.Truncate((decimal)leftValue / rightValue));
        }
    }
}
