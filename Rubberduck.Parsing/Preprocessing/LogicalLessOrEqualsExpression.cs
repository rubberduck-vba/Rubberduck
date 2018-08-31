namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class LogicalLessOrEqualsExpression : Expression
    {
        private readonly IExpression _left;
        private readonly IExpression _right;

        public LogicalLessOrEqualsExpression(IExpression left, IExpression right)
        {
            _left = left;
            _right = right;
        }

        public override IValue Evaluate()
        {
            var lt = new LogicalLessThanExpression(_left, _right).Evaluate();
            var eq = new LogicalEqualsExpression(_left, _right).Evaluate();
            if (lt == null || eq == null)
            {
                return null;
            }
            return new BoolValue(lt.AsBool || eq.AsBool);
        }
    }
}
