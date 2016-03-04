namespace Rubberduck.Parsing.Preprocessing
{
    public sealed class LogicalNotEqualsExpression : Expression
    {
        private readonly IExpression _left;
        private readonly IExpression _right;

        public LogicalNotEqualsExpression(IExpression left, IExpression right)
        {
            _left = left;
            _right = right;
        }

        public override IValue Evaluate()
        {
            var eq = new LogicalEqualsExpression(_left, _right).Evaluate();
            if (eq == null)
            {
                return null;
            }
            return new BoolValue(!eq.AsBool);
        }
    }
}
