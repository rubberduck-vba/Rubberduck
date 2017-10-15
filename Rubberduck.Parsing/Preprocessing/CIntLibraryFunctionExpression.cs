namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class CIntLibraryFunctionExpression : Expression
    {
        private readonly IExpression _expression;

        public CIntLibraryFunctionExpression(IExpression expression)
        {
            _expression = expression;
        }

        public override IValue Evaluate()
        {
            return new CCurLibraryFunctionExpression(_expression).Evaluate();
        }
    }
}
