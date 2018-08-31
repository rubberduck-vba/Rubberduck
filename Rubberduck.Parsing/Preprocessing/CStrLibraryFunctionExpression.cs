namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class CStrLibraryFunctionExpression : Expression
    {
        private readonly IExpression _expression;

        public CStrLibraryFunctionExpression(IExpression expression)
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
            return new StringValue(expr.AsString);
        }
    }
}
