using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA.DeclarationCaching;

namespace Rubberduck.Parsing.Binding
{
    public sealed class TypeBindingContext : BindingContextBase
    {
        private readonly DeclarationFinder _declarationFinder;

        public TypeBindingContext(DeclarationFinder declarationFinder)
        {
            _declarationFinder = declarationFinder;
        }

        public override IBoundExpression Resolve(
            Declaration module, 
            Declaration parent, ParserRuleContext expression,
            IBoundExpression withBlockVariable, 
            StatementResolutionContext statementContext,
            bool requiresLetCoercion = false, 
            bool isLetAssignment = false)
        {
            IExpressionBinding bindingTree = BuildTree(module, parent, expression, withBlockVariable, statementContext);
            return bindingTree?.Resolve();
        }

        public override IExpressionBinding BuildTree(
            Declaration module, 
            Declaration parent,
            ParserRuleContext expression, 
            IBoundExpression withBlockVariable,
            StatementResolutionContext statementContext, 
            bool requiresLetCoercion = false, 
            bool isLetAssignment = false)
        {
            switch (expression)
            {
                case VBAParser.LExprContext lExprContext:
                    return Visit(module, parent, lExprContext.lExpression());
                case VBAParser.CtLExprContext ctLExprContext:
                    return Visit(module, parent, ctLExprContext.lExpression());
                case VBAParser.BuiltInTypeExprContext builtInTypeExprContext:
                    return Visit(builtInTypeExprContext.builtInType());
                default:
                    return HandleUnexpectedExpressionType(expression);
            }
        }

        private IExpressionBinding Visit(VBAParser.BuiltInTypeContext builtInType)
        {
            return new BuiltInTypeDefaultBinding(builtInType);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.LExpressionContext expression)
        {
            switch (expression)
            {
                case VBAParser.SimpleNameExprContext simpleNameExprContext:
                    return Visit(module, parent, simpleNameExprContext);
                case VBAParser.MemberAccessExprContext memberAccessExprContext:
                    return Visit(module, parent, memberAccessExprContext);
                case VBAParser.IndexExprContext indexExprContext:
                    return Visit(module, parent, indexExprContext);
                default:
                    return HandleUnexpectedExpressionType(expression);
            }
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.IndexExprContext expression)
        {
            var lexpr = expression.lExpression();
            return Visit(module, parent, lexpr);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.SimpleNameExprContext expression)
        {
            return new SimpleNameTypeBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.MemberAccessExprContext expression)
        {
            var lExpression = expression.lExpression();
            var lExpressionBinding = Visit(module, parent, lExpression);
            SetPreferProjectOverUdt(lExpressionBinding);
            return new MemberAccessTypeBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression, expression.unrestrictedIdentifier(), lExpressionBinding);
        }

        private void SetPreferProjectOverUdt(IExpressionBinding lExpression)
        {
            if (!(lExpression is MemberAccessTypeBinding))
            {
                return;
            }
            var simpleNameBinding = (SimpleNameTypeBinding)((MemberAccessTypeBinding)lExpression).LExpressionBinding;
            simpleNameBinding.PreferProjectOverUdt = true;
        }
    }
}
