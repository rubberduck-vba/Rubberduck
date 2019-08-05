using Antlr4.Runtime;

namespace Rubberduck.Parsing.Binding
{
    public sealed class BuiltInTypeDefaultBinding : IExpressionBinding
    {
        private readonly ParserRuleContext _context;

        public BuiltInTypeDefaultBinding(ParserRuleContext context)
        {
            _context = context;
        }

        public IBoundExpression Resolve()
        {
            return new BuiltInTypeExpression(_context);
        }
    }
}