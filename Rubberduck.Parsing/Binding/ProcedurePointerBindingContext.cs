using System;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA.DeclarationCaching;

namespace Rubberduck.Parsing.Binding
{
    public sealed class ProcedurePointerBindingContext : IBindingContext
    {
        private readonly DeclarationFinder _declarationFinder;

        public ProcedurePointerBindingContext(DeclarationFinder declarationFinder)
        {
            _declarationFinder = declarationFinder;
        }

        public IBoundExpression Resolve(Declaration module, Declaration parent, IParseTree expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext, bool requiresLetCoercion = false, bool isLetAssignment = false)
        {
            IExpressionBinding bindingTree = BuildTree(module, parent, expression, withBlockVariable, statementContext);
            if (bindingTree != null)
            {
                return bindingTree.Resolve();
            }
            return null;
        }

        public IExpressionBinding BuildTree(Declaration module, Declaration parent, IParseTree expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext, bool requiresLetCoercion = false, bool isLetAssignment = false)
        {
            switch (expression)
            {
                case VBAParser.LExpressionContext lExpressionContext:
                    return Visit(module, parent, lExpressionContext);
                case VBAParser.ExpressionContext expressionContext:
                    return Visit(module, parent, expressionContext);
                case VBAParser.AddressOfExpressionContext addressOfExpressionContext:
                    return Visit(module, parent, addressOfExpressionContext);
                default:
                    throw new NotSupportedException($"Unexpected context type {expression.GetType()}");
            }
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.AddressOfExpressionContext expression)
        {
            return Visit(module, parent, expression.expression());
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.ExpressionContext expression)
        {
            switch (expression)
            {
                case VBAParser.LExprContext lExprContext:
                    return Visit(module, parent, lExprContext.lExpression());
                default:
                    throw new NotSupportedException($"Unexpected expression type {expression.GetType()}");
            }
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.LExpressionContext expression)
        {
            switch (expression)
            {
                case VBAParser.SimpleNameExprContext simpleNameExprContext:
                    return Visit(module, parent, simpleNameExprContext);
                case VBAParser.MemberAccessExprContext memberAccessExprContext:
                    return Visit(module, parent, memberAccessExprContext);
                default:
                    throw new NotSupportedException($"Unexpected lExpression type {expression.GetType()}");
            }
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.SimpleNameExprContext expression)
        {
            return new SimpleNameProcedurePointerBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.MemberAccessExprContext expression)
        {
            var lExpression = expression.lExpression();
            var lExpressionBinding = Visit(module, parent, lExpression);
            return new MemberAccessProcedurePointerBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression, expression.unrestrictedIdentifier(), lExpressionBinding);
        }
    }
}
