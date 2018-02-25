using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public sealed class NewTypeBinding : IExpressionBinding
    {
        private readonly DeclarationFinder _declarationFinder;
        private readonly Declaration _project;
        private readonly Declaration _module;
        private readonly Declaration _parent;
        private readonly ParserRuleContext _expression;
        private readonly IExpressionBinding _typeExpressionBinding;

        public NewTypeBinding(
            DeclarationFinder declarationFinder,
            Declaration module,
            Declaration parent,
            ParserRuleContext expression,
            IExpressionBinding typeExpressionBinding)
        {
            _declarationFinder = declarationFinder;
            _project = module.ParentDeclaration;
            _module = module;
            _parent = parent;
            _expression = expression;
            _typeExpressionBinding = typeExpressionBinding;
        }

        public IBoundExpression Resolve()
        {
            var typeExpression = _typeExpressionBinding.Resolve();
            if (typeExpression.Classification == ExpressionClassification.ResolutionFailed)
            {
                return typeExpression;
            }
            return new NewExpression(typeExpression.ReferencedDeclaration, _expression, typeExpression);
        }
    }
}
