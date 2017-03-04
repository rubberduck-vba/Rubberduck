using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public sealed class TypeBindingContext : IBindingContext
    {
        private readonly DeclarationFinder _declarationFinder;

        public TypeBindingContext(DeclarationFinder declarationFinder)
        {
            _declarationFinder = declarationFinder;
        }

        public IBoundExpression Resolve(Declaration module, Declaration parent, ParserRuleContext expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            IExpressionBinding bindingTree = BuildTree(module, parent, expression, withBlockVariable, statementContext);
            if (bindingTree != null)
            {
                return bindingTree.Resolve();
            }
            return null;
        }

        public IExpressionBinding BuildTree(Declaration module, Declaration parent, ParserRuleContext expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            dynamic dynamicExpression = expression;
            var type = expression.GetType();
            return Visit(module, parent, dynamicExpression);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.LExprContext expression)
        {
            dynamic lexpr = expression.lExpression();
            var type = expression.lExpression().GetType();
            return Visit(module, parent, lexpr);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.CtLExprContext expression)
        {
            dynamic lexpr = expression.lExpression();
            var type = expression.lExpression().GetType();
            return Visit(module, parent, lexpr);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.IndexExprContext expression)
        {
            dynamic lexpr = expression.lExpression();
            var type = expression.lExpression().GetType();
            return Visit(module, parent, lexpr);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.SimpleNameExprContext expression)
        {
            return new SimpleNameTypeBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.MemberAccessExprContext expression)
        {
            dynamic lExpression = expression.lExpression();
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
