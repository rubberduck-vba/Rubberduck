namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class CVarLibraryFunctionExpression : Expression
    {
        private readonly IExpression _expression;

        public CVarLibraryFunctionExpression(IExpression expression)
        {
            _expression = expression;
        }

        public override IValue Evaluate()
        {
            var expr = _expression.Evaluate();
            return expr;
        }
    }
}
