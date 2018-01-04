namespace Rubberduck.Parsing.PreProcessing
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
            return eq == null ? null : new BoolValue(!eq.AsBool);
        }
    }
}
