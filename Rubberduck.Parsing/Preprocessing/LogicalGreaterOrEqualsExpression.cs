namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class LogicalGreaterOrEqualsExpression : Expression
    {
        private readonly IExpression _left;
        private readonly IExpression _right;

        public LogicalGreaterOrEqualsExpression(IExpression left, IExpression right)
        {
            _left = left;
            _right = right;
        }

        public override IValue Evaluate()
        {
            var result = new LogicalLessThanExpression(_left, _right).Evaluate();
            return result == null ? null : new BoolValue(!result.AsBool);
        }
    }
}
