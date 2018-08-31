namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class ConcatExpression : Expression
    {
        private readonly IExpression _left;
        private readonly IExpression _right;

        public ConcatExpression(IExpression left, IExpression right)
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
            string leftValue = string.Empty;
            if (left != null)
            {
                leftValue = left.AsString;
            }
            string rightValue = string.Empty;
            if (right != null)
            {
                rightValue = right.AsString;
            }
            return new StringValue(leftValue + rightValue);
        }
    }
}
