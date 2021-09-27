using Antlr4.Runtime;

namespace Rubberduck.Parsing.Binding
{
    public sealed class FailedExpressionBinding : IExpressionBinding
    {
        private readonly ParserRuleContext _context;

        public FailedExpressionBinding(ParserRuleContext context)
        {
            _context = context;
        }

        public IBoundExpression Resolve()
        {
            return new ResolutionFailedExpression(_context);
        }
    }
}