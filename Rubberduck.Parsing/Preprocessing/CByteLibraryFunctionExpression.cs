namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class CByteLibraryFunctionExpression : Expression
    {
        private readonly IExpression _expression;

        public CByteLibraryFunctionExpression(IExpression expression)
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
            return new ByteValue(expr.AsByte);
        }
    }
}
