using Antlr4.Runtime;
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

        public IBoundExpression Resolve(Declaration module, Declaration parent, ParserRuleContext expression)
        {
            dynamic dynamicExpression = expression;
            IExpressionBinding bindingTree = Visit(module, parent, dynamicExpression);
            if (bindingTree != null)
            {
                return bindingTree.Resolve();
            }
            return null;
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
            return new MemberAccessTypeBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression, lExpressionBinding);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.MemberAccessExpressionContext expression)
        {
            dynamic lExpression = expression.lExpression();
            var lExpressionBinding = Visit(module, parent, lExpression);
            return new MemberAccessTypeBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression, lExpressionBinding);
        }
    }
}
