namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class CDblLibraryFunctionExpression : Expression
    {
        private readonly IExpression _expression;

        public CDblLibraryFunctionExpression(IExpression expression)
        {
            _expression = expression;
        }

        public override IValue Evaluate()
        {
            return new CCurLibraryFunctionExpression(_expression).Evaluate();
        }
    }
}
