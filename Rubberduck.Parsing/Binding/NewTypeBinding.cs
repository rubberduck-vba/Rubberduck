using Antlr4.Runtime;

namespace Rubberduck.Parsing.Binding
{
    public sealed class NewTypeBinding : IExpressionBinding
    {
        private readonly ParserRuleContext _expression;
        private readonly IExpressionBinding _typeExpressionBinding;

        public NewTypeBinding(ParserRuleContext expression,
            IExpressionBinding typeExpressionBinding)
        {
            _expression = expression;
            _typeExpressionBinding = typeExpressionBinding;
        }

        public IBoundExpression Resolve()
        {
            var typeExpression = _typeExpressionBinding.Resolve();
            return typeExpression.Classification == ExpressionClassification.ResolutionFailed 
                ? typeExpression 
                : new NewExpression(typeExpression.ReferencedDeclaration, _expression, typeExpression);
        }
    }
}
