namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class ConstantExpression : Expression
    {
        private readonly IValue _value;

        public ConstantExpression(IValue value)
        {
            _value = value;
        }

        public override IValue Evaluate()
        {
            return _value;
        }
    }
}
