namespace Rubberduck.Parsing.Preprocessing
{
    public sealed class LogicalLessOrEqualsExpression : Expression
    {
        private readonly IExpression _left;
        private readonly IExpression _right;
        private readonly VBAOptionCompare _optionCompare;

        public LogicalLessOrEqualsExpression(IExpression left, IExpression right, VBAOptionCompare optionCompare)
        {
            _left = left;
            _right = right;
            _optionCompare = optionCompare;
        }

        public override IValue Evaluate()
        {
            var lt = new LogicalLessThanExpression(_left, _right, _optionCompare).Evaluate();
            var eq = new LogicalEqualsExpression(_left, _right, _optionCompare).Evaluate();
            if (lt == null || eq == null)
            {
                return null;
            }
            return new BoolValue(lt.AsBool || eq.AsBool);
        }
    }
}
