using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public sealed class BindingService
    {
        private readonly DeclarationFinder _declarationFinder;
        private readonly IBindingContext _defaultBindingContext;
        private readonly IBindingContext _typedBindingContext;
        private readonly IBindingContext _procedurePointerBindingContext;

        public BindingService(
            DeclarationFinder declarationFinder,
            IBindingContext defaultBindingContext,
            IBindingContext typedBindingContext,
            IBindingContext procedurePointerBindingContext)
        {
            _declarationFinder = declarationFinder;
            _defaultBindingContext = defaultBindingContext;
            _typedBindingContext = typedBindingContext;
            _procedurePointerBindingContext = procedurePointerBindingContext;
        }

        public Declaration ResolveEvent(Declaration module, string identifier)
        {
            return _declarationFinder.FindEvent(module, identifier);
        }

        public Declaration ResolveGoTo(Declaration procedure, string label)
        {
            return _declarationFinder.FindLabel(procedure, label);
        }

        public IBoundExpression ResolveDefault(Declaration module, Declaration parent, ParserRuleContext expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            return _defaultBindingContext.Resolve(module, parent, expression, withBlockVariable, statementContext);
        }

        public IBoundExpression ResolveType(Declaration module, Declaration parent, ParserRuleContext expression)
        {
            var context = expression;
            var opContext = expression as VBAParser.RelationalOpContext;
            if (opContext != null && opContext.Parent is VBAParser.ComplexTypeContext)
            {
                context = opContext.GetChild<VBAParser.LExprContext>(0);
            }
            return _typedBindingContext.Resolve(module, parent, context, null, StatementResolutionContext.Undefined);
        }

        public IBoundExpression ResolveProcedurePointer(Declaration module, Declaration parent, ParserRuleContext expression)
        {
            return _procedurePointerBindingContext.Resolve(module, parent, expression, null, StatementResolutionContext.Undefined);
        }
    }
}
