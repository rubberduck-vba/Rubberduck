using Antlr4.Runtime;
using System.Collections.Generic;

namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class ConditionalCompilationBlockExpression : Expression
    {
        private readonly IEnumerable<IExpression> _children;

        public ConditionalCompilationBlockExpression(IEnumerable<IExpression> children)
        {
            _children = children;
        }

        public override IValue Evaluate()
        {
            //For some reason, using LINQ here breaks a large number of tests.
            var tokens = new List<IToken>();
            foreach(var child in _children)
            {
                tokens.AddRange(child.Evaluate().AsTokens);
            }
            return new TokensValue(tokens);
        }
    }
}
