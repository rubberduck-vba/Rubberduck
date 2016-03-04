namespace Rubberduck.Parsing.Preprocessing
{
    public sealed class LogicalGreaterOrEqualsExpression : Expression
    {
        private readonly IExpression _left;
        private readonly IExpression _right;
        private readonly VBAOptionCompare _optionCompare;

        public LogicalGreaterOrEqualsExpression(IExpression left, IExpression right, VBAOptionCompare optionCompare)
        {
            _left = left;
            _right = right;
            _optionCompare = optionCompare;
        }

        public override IValue Evaluate()
        {
            var result = new LogicalLessThanExpression(_left, _right, _optionCompare).Evaluate();
            if (result == null)
            {
                return null;
            }
            return new BoolValue(!result.AsBool);
        }
    }
}
