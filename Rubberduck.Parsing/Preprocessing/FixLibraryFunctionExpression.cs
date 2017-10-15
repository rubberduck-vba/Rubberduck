namespace Rubberduck.Parsing.Preprocessing
{
    public sealed class FixLibraryFunctionExpression : Expression
    {
        private readonly IExpression _expression;

        public FixLibraryFunctionExpression(IExpression expression)
        {
            _expression = expression;
        }

        public override IValue Evaluate()
        {
            return new IntLibraryFunctionExpression(_expression).Evaluate();
        }
    }
}
