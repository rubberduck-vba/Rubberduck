using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public sealed class InstanceDefaultBinding : IExpressionBinding
    {
        private readonly ParserRuleContext _context;
        private readonly Declaration _module;

        public InstanceDefaultBinding(
            ParserRuleContext context,
            Declaration module)
        {
            _context = context;
            _module = module;
        }

        public IBoundExpression Resolve()
        {
            return new InstanceExpression(_module, _context);
        }
    }
}
