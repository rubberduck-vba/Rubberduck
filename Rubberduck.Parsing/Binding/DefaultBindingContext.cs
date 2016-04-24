using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public sealed class DefaultBindingContext : IBindingContext
    {
        private readonly DeclarationFinder _declarationFinder;

        public DefaultBindingContext(DeclarationFinder declarationFinder)
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

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.NewExprContext expression)
        {
            return Visit(module, parent, expression.newExpression());
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.NewExpressionContext expression)
        {
            var typeExpressionBinding = Visit(module, parent, expression.typeExpression());
            if (typeExpressionBinding == null)
            {
                return null;
            }
            return new NewTypeBinding(_declarationFinder, module, parent, expression, typeExpressionBinding);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.TypeExpressionContext expression)
        {
            if (expression.builtInType() != null)
            {
                return null;
            }
            return Visit(module, parent, expression.definedTypeExpression());
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.DefinedTypeExpressionContext expression)
        {
            if (expression.simpleNameExpression() != null)
            {
                return Visit(module, parent, expression.simpleNameExpression());
            }
            return Visit(module, parent, expression.memberAccessExpression());
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.SimpleNameExprContext expression)
        {
            return Visit(module, parent, expression.simpleNameExpression());
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.SimpleNameExpressionContext expression)
        {
            return new SimpleNameTypeBinding(_declarationFinder, module, parent, expression);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.MemberAccessExprContext expression)
        {
            dynamic lExpression = expression.lExpression();
            var lExpressionBinding = Visit(module, parent, lExpression);
            return new MemberAccessTypeBinding(_declarationFinder, module, parent, expression, lExpressionBinding);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.MemberAccessExpressionContext expression)
        {
            dynamic lExpression = expression.lExpression();
            var lExpressionBinding = Visit(module, parent, lExpression);
            return new MemberAccessTypeBinding(_declarationFinder, module, parent, expression, lExpressionBinding);
        }
    }
}
