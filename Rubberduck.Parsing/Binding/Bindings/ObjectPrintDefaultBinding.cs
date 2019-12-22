using Antlr4.Runtime;

namespace Rubberduck.Parsing.Binding
{
    public sealed class ObjectPrintDefaultBinding : IExpressionBinding
    {
        private readonly ParserRuleContext _context;
        private readonly IExpressionBinding _printMethodBinding;
        private readonly IExpressionBinding _outputListBinding;

        public ObjectPrintDefaultBinding(
            ParserRuleContext context,
            IExpressionBinding printMethodBinding,
            IExpressionBinding outputListBinding)
        {
            _context = context;
            _printMethodBinding = printMethodBinding;
            _outputListBinding = outputListBinding;
        }

        public IBoundExpression Resolve()
        {
            var printMethodExpression = _printMethodBinding.Resolve();
            var outputListExpression = _outputListBinding?.Resolve();
            return new ObjectPrintExpression(_context, printMethodExpression, outputListExpression);
        }
    }
}