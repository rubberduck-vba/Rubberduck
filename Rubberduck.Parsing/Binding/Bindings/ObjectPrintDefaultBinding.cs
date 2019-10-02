using Antlr4.Runtime;

namespace Rubberduck.Parsing.Binding
{
    public sealed class ObjectPrintDefaultBinding : IExpressionBinding
    {
        private readonly ParserRuleContext _context;
        private readonly IExpressionBinding _memberAccessBinding;
        private readonly IExpressionBinding _outputListBinding;

        public ObjectPrintDefaultBinding(
            ParserRuleContext context,
            IExpressionBinding memberAccessBinding,
            IExpressionBinding outputListBinding)
        {
            _context = context;
            _memberAccessBinding = memberAccessBinding;
            _outputListBinding = outputListBinding;
        }

        public IBoundExpression Resolve()
        {
            var memberAccessExpression = _memberAccessBinding.Resolve();
            var outputListExpression = _outputListBinding.Resolve();
            return new ObjectPrintExpression(_context, memberAccessExpression, outputListExpression);
        }
    }
}