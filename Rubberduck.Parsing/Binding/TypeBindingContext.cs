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

        public IBoundExpression Resolve(Declaration module, Declaration parent, ParserRuleContext expression)
        {
            var lexpr = ((VBAExpressionParser.LExprContext)expression).lExpression();
            IExpressionBinding bindingTree = null;
            if (lexpr is VBAExpressionParser.SimpleNameExprContext)
            {
                bindingTree = ResolveSimpleNameExpression(module, parent, ((VBAExpressionParser.SimpleNameExprContext)lexpr).simpleNameExpression());
            }
            if (bindingTree != null)
            {
                return bindingTree.Resolve();
            }
            return null;
        }

        private IExpressionBinding ResolveSimpleNameExpression(Declaration module, Declaration parent, VBAExpressionParser.SimpleNameExpressionContext expression)
        {
            return new SimpleNameTypeBinding(_declarationFinder, module, expression);
        }
    }
}
