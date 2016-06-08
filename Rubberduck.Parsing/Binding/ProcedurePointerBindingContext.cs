using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public sealed class ProcedurePointerBindingContext : IBindingContext
    {
        private readonly DeclarationFinder _declarationFinder;

        public ProcedurePointerBindingContext(DeclarationFinder declarationFinder)
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
            return Visit(module, parent, dynamicExpression);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.ExpressionContext expression)
        {
            return Visit(module, parent, (dynamic)expression);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.AddressOfExpressionContext expression)
        {
            return Visit(module, parent, (dynamic)expression.expression());
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.LExprContext expression)
        {
            dynamic lexpr = expression.lExpression();
            return Visit(module, parent, lexpr);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.SimpleNameExprContext expression)
        {
            return new SimpleNameProcedurePointerBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.MemberAccessExprContext expression)
        {
            dynamic lExpression = expression.lExpression();
            var lExpressionBinding = Visit(module, parent, lExpression);
            return new MemberAccessProcedurePointerBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression, expression.unrestrictedIdentifier(), lExpressionBinding);
        }
    }
}
