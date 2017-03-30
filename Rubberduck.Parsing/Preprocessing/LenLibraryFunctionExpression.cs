namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class LenLibraryFunctionExpression : Expression
    {
        private readonly IExpression _expression;

        public LenLibraryFunctionExpression(IExpression expression)
        {
            _expression = expression;
        }

        public override IValue Evaluate()
        {
            var expr = _expression.Evaluate();
            if (expr == null)
            {
                return null;
            }
            return new DecimalValue(expr.AsString.Length);
        }
    }
}
