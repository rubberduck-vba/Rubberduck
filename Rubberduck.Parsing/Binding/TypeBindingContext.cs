using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;
using System;

namespace Rubberduck.Parsing.Binding
{
    public sealed class TypeBindingContext : IBindingContext
    {
        private readonly DeclarationFinder _declarationFinder;

        public TypeBindingContext(DeclarationFinder declarationFinder)
        {
            _declarationFinder = declarationFinder;
        }

        public IBoundExpression Resolve(Declaration module, Declaration parent, ParserRuleContext expression, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            IExpressionBinding bindingTree = BuildTree(module, parent, expression, withBlockVariable, statementContext);
            if (bindingTree != null)
            {
                return bindingTree.Resolve();
            }
            return null;
        }

        public IExpressionBinding BuildTree(Declaration module, Declaration parent, ParserRuleContext expression, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            dynamic dynamicExpression = expression;
            return Visit(module, parent, dynamicExpression);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.StartRuleContext expression)
        {
            return Visit(module, parent, (dynamic)expression.expression());
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.LExprContext expression)
        {
            dynamic lexpr = expression.lExpression();
            return Visit(module, parent, lexpr);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.SimpleNameExprContext expression)
        {
            var simpleNameExpression = expression.simpleNameExpression();
            return Visit(module, parent, simpleNameExpression);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.SimpleNameExpressionContext expression)
        {
            return new SimpleNameTypeBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.MemberAccessExprContext expression)
        {
            dynamic lExpression = expression.lExpression();
            var lExpressionBinding = Visit(module, parent, lExpression);
            SetPreferProjectOverUdt(lExpressionBinding);
            return new MemberAccessTypeBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression, expression.unrestrictedName(), lExpressionBinding);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.MemberAccessExpressionContext expression)
        {
            dynamic lExpression = expression.lExpression();
            var lExpressionBinding = Visit(module, parent, lExpression);
            SetPreferProjectOverUdt(lExpressionBinding);
            return new MemberAccessTypeBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression, expression.unrestrictedName() , lExpressionBinding);
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
