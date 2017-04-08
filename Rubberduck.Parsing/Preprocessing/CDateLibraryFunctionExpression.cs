﻿namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class CDateLibraryFunctionExpression : Expression
    {
        private readonly IExpression _expression;

        public CDateLibraryFunctionExpression(IExpression expression)
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
            return new DateValue(expr.AsDate);
        }
    }
}
