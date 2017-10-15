namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class IsExpression : Expression
    {
        private readonly IExpression _left;
        private readonly IExpression _right;

        public IsExpression(IExpression left, IExpression right)
        {
            _left = left;
            _right = right;
        }

        public override IValue Evaluate()
        {
            var left = _left.Evaluate();
            var right = _right.Evaluate();
            return new BoolValue(left == null && right == null);
        }
    }
}
