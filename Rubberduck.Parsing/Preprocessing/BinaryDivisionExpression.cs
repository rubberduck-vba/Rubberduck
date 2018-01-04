namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class BinaryDivisionExpression : Expression
    {
        private readonly IExpression _left;
        private readonly IExpression _right;

        public BinaryDivisionExpression(IExpression left, IExpression right)
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

            return new DecimalValue(left.AsDecimal / right.AsDecimal);
        }
    }
}
