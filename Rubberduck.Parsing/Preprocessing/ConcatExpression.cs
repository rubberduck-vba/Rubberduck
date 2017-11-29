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

            var leftValue = left == null ? string.Empty : left.AsString;

            var rightValue = right == null ? string.Empty : right.AsString;

            return new StringValue(leftValue + rightValue);
        }
    }
}
