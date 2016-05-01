using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public sealed class DefaultBindingContext : IBindingContext
    {
        private readonly DeclarationFinder _declarationFinder;
        private readonly IBindingContext _typeBindingContext;
        private readonly IBindingContext _procedurePointerBindingContext;

        public DefaultBindingContext(
            DeclarationFinder declarationFinder,
            IBindingContext typeBindingContext,
            IBindingContext procedurePointerBindingContext)
        {
            _declarationFinder = declarationFinder;
            _typeBindingContext = typeBindingContext;
            _procedurePointerBindingContext = procedurePointerBindingContext;
        }

        public IBoundExpression Resolve(Declaration module, Declaration parent, ParserRuleContext expression, IBoundExpression withBlockVariable)
        {
            IExpressionBinding bindingTree = BuildTree(module, parent, expression, withBlockVariable);
            if (bindingTree != null)
            {
                return bindingTree.Resolve();
            }
            return null;
        }

        public IExpressionBinding BuildTree(Declaration module, Declaration parent, ParserRuleContext expression, IBoundExpression withBlockVariable)
        {
            dynamic dynamicExpression = expression;
            return Visit(module, parent, dynamicExpression, withBlockVariable);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.LExprContext expression, IBoundExpression withBlockVariable)
        {
            dynamic lexpr = expression.lExpression();
            return Visit(module, parent, lexpr, withBlockVariable);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.NewExprContext expression, IBoundExpression withBlockVariable)
        {
            return Visit(module, parent, expression.newExpression(), withBlockVariable);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.NewExpressionContext expression, IBoundExpression withBlockVariable)
        {
            var typeExpressionBinding = Visit(module, parent, expression.typeExpression(), withBlockVariable);
            if (typeExpressionBinding == null)
            {
                return null;
            }
            return new NewTypeBinding(_declarationFinder, module, parent, expression, typeExpressionBinding);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.TypeExpressionContext expression, IBoundExpression withBlockVariable)
        {
            if (expression.builtInType() != null)
            {
                return null;
            }
            return Visit(module, parent, expression.definedTypeExpression(), withBlockVariable);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.DefinedTypeExpressionContext expression, IBoundExpression withBlockVariable)
        {
            if (expression.simpleNameExpression() != null)
            {
                return _typeBindingContext.BuildTree(module, parent, expression.simpleNameExpression(), withBlockVariable);
            }
            return _typeBindingContext.BuildTree(module, parent, expression.memberAccessExpression(), withBlockVariable);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.SimpleNameExprContext expression, IBoundExpression withBlockVariable)
        {
            return Visit(module, parent, expression.simpleNameExpression(), withBlockVariable);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.SimpleNameExpressionContext expression, IBoundExpression withBlockVariable)
        {
            return new SimpleNameDefaultBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.MemberAccessExprContext expression, IBoundExpression withBlockVariable)
        {
            dynamic lExpression = expression.lExpression();
            var lExpressionBinding = Visit(module, parent, lExpression, withBlockVariable);
            return new MemberAccessDefaultBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression, lExpressionBinding);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.MemberAccessExpressionContext expression, IBoundExpression withBlockVariable)
        {
            dynamic lExpression = expression.lExpression();
            var lExpressionBinding = Visit(module, parent, lExpression, withBlockVariable);
            return new MemberAccessDefaultBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression, lExpressionBinding);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.IndexExprContext expression, IBoundExpression withBlockVariable)
        {
            dynamic lExpression = expression.lExpression();
            var lExpressionBinding = Visit(module, parent, lExpression, withBlockVariable);
            var argumentListBinding = VisitArgumentList(module, parent, expression.argumentList(), withBlockVariable);
            return new IndexDefaultBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression, lExpressionBinding, argumentListBinding);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.IndexExpressionContext expression, IBoundExpression withBlockVariable)
        {
            dynamic lExpression = expression.lExpression();
            var lExpressionBinding = Visit(module, parent, lExpression, withBlockVariable);
            var argumentListBinding = VisitArgumentList(module, parent, expression.argumentList(), withBlockVariable);
            return new IndexDefaultBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression, lExpressionBinding, argumentListBinding);
        }

        private ArgumentList VisitArgumentList(Declaration module, Declaration parent, VBAExpressionParser.ArgumentListContext argumentList, IBoundExpression withBlockVariable)
        {
            var convertedList = new ArgumentList();
            var list = argumentList.positionalOrNamedArgumentList();
            if (list.positionalArgument() != null)
            {
                foreach (var expr in list.positionalArgument())
                {
                    convertedList.AddArgument(VisitArgumentBinding(module, parent, expr.argumentExpression(), withBlockVariable), ArgumentListArgumentType.Positional);
                }
            }
            if (list.requiredPositionalArgument() != null)
            {
                convertedList.AddArgument(VisitArgumentBinding(module, parent, list.requiredPositionalArgument().argumentExpression(), withBlockVariable), ArgumentListArgumentType.Positional);
            }
            if (list.namedArgumentList() != null)
            {
                foreach (var expr in list.namedArgumentList().namedArgument())
                {
                    convertedList.AddArgument(VisitArgumentBinding(module, parent, expr.argumentExpression(), withBlockVariable), ArgumentListArgumentType.Named);
                }
            }
            return convertedList;
        }

        private IExpressionBinding VisitArgumentBinding(Declaration module, Declaration parent, VBAExpressionParser.ArgumentExpressionContext argumentExpression, IBoundExpression withBlockVariable)
        {
            if (argumentExpression.expression() != null)
            {
                dynamic expr = argumentExpression.expression();
                return Visit(module, parent, expr, withBlockVariable);
            }
            else
            {
                dynamic expr = argumentExpression.addressOfExpression();
                return Visit(module, parent, expr, withBlockVariable);
            }
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.DictionaryAccessExprContext expression, IBoundExpression withBlockVariable)
        {
            dynamic lExpression = expression.lExpression();
            var lExpressionBinding = Visit(module, parent, lExpression, withBlockVariable);
            return VisitDictionaryAccessExpression(module, parent, expression, expression.unrestrictedName(), lExpressionBinding);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.DictionaryAccessExpressionContext expression, IBoundExpression withBlockVariable)
        {
            dynamic lExpression = expression.lExpression();
            var lExpressionBinding = Visit(module, parent, lExpression, withBlockVariable);
            return VisitDictionaryAccessExpression(module, parent, expression, expression.unrestrictedName(), lExpressionBinding);
        }

        private IExpressionBinding VisitDictionaryAccessExpression(Declaration module, Declaration parent, ParserRuleContext expression, ParserRuleContext nameContext, IExpressionBinding lExpressionBinding)
        {
            /*
                A dictionary access expression is syntactically translated into an index expression with the same 
                expression for <l-expression> and an argument list with a single positional argument with a 
                declared type of String and a value equal to the name value of <unrestricted-name>. 
             */
            var fakeArgList = new ArgumentList();
            fakeArgList.AddArgument(new LiteralDefaultBinding(nameContext), ArgumentListArgumentType.Positional);
            return new IndexDefaultBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression, lExpressionBinding, fakeArgList);
        }

        private IExpressionBinding VisitDictionaryAccessExpression(Declaration module, Declaration parent, ParserRuleContext expression, ParserRuleContext nameContext, IBoundExpression lExpression)
        {
            /*
                A dictionary access expression is syntactically translated into an index expression with the same 
                expression for <l-expression> and an argument list with a single positional argument with a 
                declared type of String and a value equal to the name value of <unrestricted-name>. 
             */
            var fakeArgList = new ArgumentList();
            fakeArgList.AddArgument(new LiteralDefaultBinding(nameContext), ArgumentListArgumentType.Positional);
            return new IndexDefaultBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression, lExpression, fakeArgList);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.WithExprContext expression, IBoundExpression withBlockVariable)
        {
            return Visit(module, parent, expression.withExpression(), withBlockVariable);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.WithExpressionContext expression, IBoundExpression withBlockVariable)
        {
            /*
                A <with-member-access-expression> or <with-dictionary-access-expression> is 
                statically resolved as a normal member access or dictionary access expression, respectively, as if 
                the innermost enclosing With block variable was specified for <l-expression>. If there is no 
                enclosing With block, the <with-expression> is invalid.
             */
            if (expression.withMemberAccessExpression() != null)
            {
                return new MemberAccessDefaultBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression, withBlockVariable, expression.withMemberAccessExpression().unrestrictedName().GetText());
            }
            else
            {
                return VisitDictionaryAccessExpression(module, parent, expression, expression.withDictionaryAccessExpression().unrestrictedName(), withBlockVariable);
            }
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.ParenthesizedExprContext expression, IBoundExpression withBlockVariable)
        {
            dynamic expressionParens = expression.expression();
            var expressionBinding = Visit(module, parent, expressionParens, withBlockVariable);
            return new ParenthesizedDefaultBinding(expression, expressionBinding);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.PowOpContext expression, IBoundExpression withBlockVariable)
        {
            return VisitBinaryOp(module, parent, expression, expression.expression()[0], expression.expression()[1], withBlockVariable);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.MultOpContext expression, IBoundExpression withBlockVariable)
        {
            return VisitBinaryOp(module, parent, expression, expression.expression()[0], expression.expression()[1], withBlockVariable);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.IntDivOpContext expression, IBoundExpression withBlockVariable)
        {
            return VisitBinaryOp(module, parent, expression, expression.expression()[0], expression.expression()[1], withBlockVariable);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.AddOpContext expression, IBoundExpression withBlockVariable)
        {
            return VisitBinaryOp(module, parent, expression, expression.expression()[0], expression.expression()[1], withBlockVariable);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.ConcatOpContext expression, IBoundExpression withBlockVariable)
        {
            return VisitBinaryOp(module, parent, expression, expression.expression()[0], expression.expression()[1], withBlockVariable);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.RelationalOpContext expression, IBoundExpression withBlockVariable)
        {
            return VisitBinaryOp(module, parent, expression, expression.expression()[0], expression.expression()[1], withBlockVariable);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.LogicalAndOpContext expression, IBoundExpression withBlockVariable)
        {
            return VisitBinaryOp(module, parent, expression, expression.expression()[0], expression.expression()[1], withBlockVariable);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.LogicalOrOpContext expression, IBoundExpression withBlockVariable)
        {
            return VisitBinaryOp(module, parent, expression, expression.expression()[0], expression.expression()[1], withBlockVariable);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.LogicalXorOpContext expression, IBoundExpression withBlockVariable)
        {
            return VisitBinaryOp(module, parent, expression, expression.expression()[0], expression.expression()[1], withBlockVariable);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.LogicalEqvOpContext expression, IBoundExpression withBlockVariable)
        {
            return VisitBinaryOp(module, parent, expression, expression.expression()[0], expression.expression()[1], withBlockVariable);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.LogicalImpOpContext expression, IBoundExpression withBlockVariable)
        {
            return VisitBinaryOp(module, parent, expression, expression.expression()[0], expression.expression()[1], withBlockVariable);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.UnaryMinusOpContext expression, IBoundExpression withBlockVariable)
        {
            return VisitUnaryOp(module, parent, expression, expression.expression(), withBlockVariable);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.LogicalNotOpContext expression, IBoundExpression withBlockVariable)
        {
            return VisitUnaryOp(module, parent, expression, expression.expression(), withBlockVariable);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.LiteralExprContext expression, IBoundExpression withBlockVariable)
        {
            return Visit(module, parent, expression.literalExpression(), withBlockVariable);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.InstanceExprContext expression, IBoundExpression withBlockVariable)
        {
            return Visit(module, parent, expression.instanceExpression(), withBlockVariable);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.InstanceExpressionContext expression, IBoundExpression withBlockVariable)
        {
            return new InstanceDefaultBinding(expression, module);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.LiteralExpressionContext expression, IBoundExpression withBlockVariable)
        {
            return new LiteralDefaultBinding(expression);
        }

        private IExpressionBinding VisitBinaryOp(Declaration module, Declaration parent, ParserRuleContext context, ParserRuleContext left, ParserRuleContext right, IBoundExpression withBlockVariable)
        {
            dynamic leftExpr = left;
            var leftBinding = Visit(module, parent, leftExpr, withBlockVariable);
            dynamic rightExpr = right;
            var rightBinding = Visit(module, parent, rightExpr, withBlockVariable);
            return new BinaryOpDefaultBinding(context, leftBinding, rightBinding);
        }

        private IExpressionBinding VisitUnaryOp(Declaration module, Declaration parent, ParserRuleContext context, ParserRuleContext expr, IBoundExpression withBlockVariable)
        {
            dynamic exprExpr = expr;
            var exprBinding = Visit(module, parent, exprExpr, withBlockVariable);
            return new UnaryOpDefaultBinding(context, exprExpr);
        }
    }
}
