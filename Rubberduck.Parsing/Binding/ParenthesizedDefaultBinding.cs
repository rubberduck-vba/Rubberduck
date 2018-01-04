using Antlr4.Runtime;

namespace Rubberduck.Parsing.Binding
{
    public sealed class ParenthesizedDefaultBinding : IExpressionBinding
    {
        private readonly ParserRuleContext _context;
        private readonly IExpressionBinding _expressionBinding;

        public ParenthesizedDefaultBinding(
            ParserRuleContext context,
            IExpressionBinding expressionBinding)
        {
            _context = context;
            _expressionBinding = expressionBinding;
        }

        public IBoundExpression Resolve()
        {
            var expr = _expressionBinding.Resolve();
            return expr.Classification == ExpressionClassification.ResolutionFailed 
                ? expr 
                : new ParenthesizedExpression(expr.ReferencedDeclaration, _context, expr);
        }
    }
}
