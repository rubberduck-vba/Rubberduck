using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public sealed class MissingArgumentBinding : IExpressionBinding
    {
        private readonly Declaration _parent;
        private readonly ParserRuleContext _missingArgumentContext;

        public MissingArgumentBinding(ParserRuleContext missingArgumentContext)
        {
            _missingArgumentContext = missingArgumentContext;
        }

        public IBoundExpression Resolve()
        {
            return new MissingArgumentExpression(ExpressionClassification.Variable, _missingArgumentContext);
        }
    }
}