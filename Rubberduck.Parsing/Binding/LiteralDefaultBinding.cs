using Antlr4.Runtime;

namespace Rubberduck.Parsing.Binding
{
    public sealed class LiteralDefaultBinding : IExpressionBinding
    {
        private readonly ParserRuleContext _context;

        public LiteralDefaultBinding(ParserRuleContext context)
        {
            _context = context;
        }

        public IBoundExpression Resolve()
        {
            return new LiteralExpression(_context);
        }
    }
}
