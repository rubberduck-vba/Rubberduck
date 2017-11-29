namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class CCurLibraryFunctionExpression : Expression
    {
        private readonly IExpression _expression;

        public CCurLibraryFunctionExpression(IExpression expression)
        {
            _expression = expression;
        }

        public override IValue Evaluate()
        {
            var expr = _expression.Evaluate();
            return expr == null ? null : new DecimalValue(expr.AsDecimal);
        }
    }
}
