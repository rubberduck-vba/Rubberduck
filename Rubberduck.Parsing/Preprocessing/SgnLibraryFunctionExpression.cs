using System;

namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class SgnLibraryFunctionExpression : Expression
    {
        private readonly IExpression _expression;

        public SgnLibraryFunctionExpression(IExpression expression)
        {
            _expression = expression;
        }

        public override IValue Evaluate()
        {
            var expr = _expression.Evaluate();
            return expr == null 
                ? null 
                : new DecimalValue(Math.Sign(expr.AsDecimal));
        }
    }
}
