using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Parsing.Preprocessing
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
            return new StringValue(string.Join(string.Empty, _children.Select(child => child.Evaluate().AsString)));
        }
    }
}
