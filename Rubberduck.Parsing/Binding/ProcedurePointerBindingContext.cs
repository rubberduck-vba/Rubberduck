using Antlr4.Runtime;
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

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.AddressOfExpressionContext expression)
        {
            return Visit(module, parent, expression.procedurePointerExpression());
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.ProcedurePointerExpressionContext expression)
        {
            if (expression.memberAccessExpression() != null)
            {
                return Visit(module, parent, expression.memberAccessExpression());
            }
            else
            {
                return Visit(module, parent, expression.simpleNameExpression());
            }
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
            return new SimpleNameProcedurePointerBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.MemberAccessExprContext expression)
        {
            dynamic lExpression = expression.lExpression();
            var lExpressionBinding = Visit(module, parent, lExpression);
            return new MemberAccessProcedurePointerBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression, lExpressionBinding);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.MemberAccessExpressionContext expression)
        {
            dynamic lExpression = expression.lExpression();
            var lExpressionBinding = Visit(module, parent, lExpression);
            return new MemberAccessProcedurePointerBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression, lExpressionBinding);
        }
    }
}
