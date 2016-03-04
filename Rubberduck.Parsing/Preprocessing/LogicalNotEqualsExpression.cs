namespace Rubberduck.Parsing.Preprocessing
{
    public sealed class LogicalNotEqualsExpression : Expression
    {
        private readonly IExpression _left;
        private readonly IExpression _right;
        private readonly VBAOptionCompare _optionCompare;

        public LogicalNotEqualsExpression(IExpression left, IExpression right, VBAOptionCompare optionCompare)
        {
            _left = left;
            _right = right;
            _optionCompare = optionCompare;
        }

        public override IValue Evaluate()
        {
            var eq = new LogicalEqualsExpression(_left, _right, _optionCompare).Evaluate();
            if (eq == null)
            {
                return null;
            }
            return new BoolValue(!eq.AsBool);
        }
    }
}
