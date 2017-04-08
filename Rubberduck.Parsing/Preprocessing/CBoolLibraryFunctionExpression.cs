namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class CBoolLibraryFunctionExpression : Expression
    {
        private readonly IExpression _expression;

        public CBoolLibraryFunctionExpression(IExpression expression)
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
            return new BoolValue(expr.AsBool);
        }
    }
}
