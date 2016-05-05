using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;
using System;

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
            return Visit(module, parent, dynamicExpression, withBlockVariable, statementContext);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.StartRuleContext expression, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            // Call statements always have an argument list
            if (statementContext == ResolutionStatementContext.CallStatement)
            {
                if (expression.callStmt() != null)
                {
                    return VisitCallStmt(module, parent, expression.callStmt(), withBlockVariable, statementContext);
                }
                else
                {
                    return VisitCallStmt(module, parent, expression.expression(), withBlockVariable, statementContext);
                }
            }
            return Visit(module, parent, (dynamic)expression.expression(), withBlockVariable, statementContext);
        }

        private IExpressionBinding VisitCallStmt(Declaration module, Declaration parent, ParserRuleContext expression, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            if (expression is VBAExpressionParser.CallStmtContext)
            {
                var callStmtExpression = (VBAExpressionParser.CallStmtContext)expression;
                dynamic lexpr;
                if (callStmtExpression.simpleNameExpression() != null)
                {
                    lexpr = callStmtExpression.simpleNameExpression();
                }
                else if (callStmtExpression.memberAccessExpression() != null)
                {
                    lexpr = callStmtExpression.memberAccessExpression();
                }
                else
                {
                    lexpr = callStmtExpression.withExpression();
                }
                var lexprBinding = Visit(module, parent, lexpr, withBlockVariable, ResolutionStatementContext.Undefined);
                var argList = VisitArgumentList(module, parent, callStmtExpression.argumentList(), withBlockVariable, ResolutionStatementContext.Undefined);
                return new IndexDefaultBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression, lexprBinding, argList);
            }
            else
            {
                var lexprBinding = Visit(module, parent, (dynamic)expression, withBlockVariable, ResolutionStatementContext.Undefined);
                if (!(lexprBinding is IndexDefaultBinding))
                {
                    return new IndexDefaultBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression, lexprBinding, new ArgumentList());
                }
                else
                {
                    return lexprBinding;
                }
            }
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.LExprContext expression, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            dynamic lexpr = expression.lExpression();
            return Visit(module, parent, lexpr, withBlockVariable, statementContext);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.NewExprContext expression, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            return Visit(module, parent, expression.newExpression(), withBlockVariable, ResolutionStatementContext.Undefined);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.NewExpressionContext expression, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            var typeExpressionBinding = Visit(module, parent, expression.typeExpression(), withBlockVariable, ResolutionStatementContext.Undefined);
            if (typeExpressionBinding == null)
            {
                return null;
            }
            return new NewTypeBinding(_declarationFinder, module, parent, expression, typeExpressionBinding);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.TypeExpressionContext expression, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            if (expression.builtInType() != null)
            {
                return null;
            }
            return Visit(module, parent, expression.definedTypeExpression(), withBlockVariable, ResolutionStatementContext.Undefined);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.DefinedTypeExpressionContext expression, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            if (expression.simpleNameExpression() != null)
            {
                return _typeBindingContext.BuildTree(module, parent, expression.simpleNameExpression(), withBlockVariable, ResolutionStatementContext.Undefined);
            }
            return _typeBindingContext.BuildTree(module, parent, expression.memberAccessExpression(), withBlockVariable, ResolutionStatementContext.Undefined);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.SimpleNameExprContext expression, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            return Visit(module, parent, expression.simpleNameExpression(), withBlockVariable, ResolutionStatementContext.Undefined);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.SimpleNameExpressionContext expression, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            return new SimpleNameDefaultBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression, statementContext);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.MemberAccessExprContext expression, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            dynamic lExpression = expression.lExpression();
            var lExpressionBinding = Visit(module, parent, lExpression, withBlockVariable, ResolutionStatementContext.Undefined);
            return new MemberAccessDefaultBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression, lExpressionBinding, statementContext, expression.unrestrictedName());
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.MemberAccessExpressionContext expression, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            dynamic lExpression = expression.lExpression();
            var lExpressionBinding = Visit(module, parent, lExpression, withBlockVariable, ResolutionStatementContext.Undefined);
            return new MemberAccessDefaultBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression, lExpressionBinding, statementContext, expression.unrestrictedName());
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.IndexExprContext expression, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            dynamic lExpression = expression.lExpression();
            var lExpressionBinding = Visit(module, parent, lExpression, withBlockVariable, ResolutionStatementContext.Undefined);
            var argumentListBinding = VisitArgumentList(module, parent, expression.argumentList(), withBlockVariable, ResolutionStatementContext.Undefined);
            return new IndexDefaultBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression, lExpressionBinding, argumentListBinding);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.IndexExpressionContext expression, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            dynamic lExpression = expression.lExpression();
            var lExpressionBinding = Visit(module, parent, lExpression, withBlockVariable, ResolutionStatementContext.Undefined);
            var argumentListBinding = VisitArgumentList(module, parent, expression.argumentList(), withBlockVariable, ResolutionStatementContext.Undefined);
            return new IndexDefaultBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression, lExpressionBinding, argumentListBinding);
        }

        private ArgumentList VisitArgumentList(Declaration module, Declaration parent, VBAExpressionParser.ArgumentListContext argumentList, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            var convertedList = new ArgumentList();
            if (argumentList == null)
            {
                return convertedList;
            }
            var list = argumentList.positionalOrNamedArgumentList();
            if (list.positionalArgument() != null)
            {
                foreach (var expr in list.positionalArgument())
                {
                    convertedList.AddArgument(new ArgumentListArgument(
                        VisitArgumentBinding(module, parent, expr.argumentExpression(), withBlockVariable,
                        ResolutionStatementContext.Undefined), ArgumentListArgumentType.Positional));
                }
            }
            if (list.requiredPositionalArgument() != null)
            {
                convertedList.AddArgument(new ArgumentListArgument(
                    VisitArgumentBinding(module, parent, list.requiredPositionalArgument().argumentExpression(),
                    withBlockVariable, ResolutionStatementContext.Undefined),
                    ArgumentListArgumentType.Positional));
            }
            if (list.namedArgumentList() != null)
            {
                foreach (var expr in list.namedArgumentList().namedArgument())
                {
                    convertedList.AddArgument(new ArgumentListArgument(
                        VisitArgumentBinding(module, parent, expr.argumentExpression(),
                        withBlockVariable,
                        ResolutionStatementContext.Undefined),
                        ArgumentListArgumentType.Named,
                        CreateNamedArgumentExpressionCreator(expr.unrestrictedName().GetText(), expr.unrestrictedName())));
                }
            }
            return convertedList;
        }

        private Func<Declaration, IBoundExpression> CreateNamedArgumentExpressionCreator(string parameterName, ParserRuleContext context)
        {
            return calledProcedure =>
            {
                ExpressionClassification classification;
                if (calledProcedure.DeclarationType == DeclarationType.Procedure)
                {
                    classification = ExpressionClassification.Subroutine;
                }
                else if (calledProcedure.DeclarationType == DeclarationType.Function || calledProcedure.DeclarationType == DeclarationType.LibraryFunction || calledProcedure.DeclarationType == DeclarationType.LibraryProcedure)
                {
                    classification = ExpressionClassification.Function;
                }
                else
                {
                    classification = ExpressionClassification.Property;
                }
                var parameter = _declarationFinder.FindParameter(calledProcedure, parameterName);
                if (parameter != null)
                {
                    return new SimpleNameExpression(parameter, classification, context);
                }
                else
                {
                    return null;
                }
            };
        }

        private IExpressionBinding VisitArgumentBinding(Declaration module, Declaration parent, VBAExpressionParser.ArgumentExpressionContext argumentExpression, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            if (argumentExpression.expression() != null)
            {
                dynamic expr = argumentExpression.expression();
                return Visit(module, parent, expr, withBlockVariable, ResolutionStatementContext.Undefined);
            }
            else
            {
                dynamic expr = argumentExpression.addressOfExpression();
                return Visit(module, parent, expr, withBlockVariable, ResolutionStatementContext.Undefined);
            }
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.AddressOfExpressionContext expression, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            return _procedurePointerBindingContext.BuildTree(module, parent, expression, withBlockVariable, statementContext);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.DictionaryAccessExprContext expression, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            dynamic lExpression = expression.lExpression();
            var lExpressionBinding = Visit(module, parent, lExpression, withBlockVariable, ResolutionStatementContext.Undefined);
            return VisitDictionaryAccessExpression(module, parent, expression, expression.unrestrictedName(), lExpressionBinding, ResolutionStatementContext.Undefined);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.DictionaryAccessExpressionContext expression, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            dynamic lExpression = expression.lExpression();
            var lExpressionBinding = Visit(module, parent, lExpression, withBlockVariable, ResolutionStatementContext.Undefined);
            return VisitDictionaryAccessExpression(module, parent, expression, expression.unrestrictedName(), lExpressionBinding, ResolutionStatementContext.Undefined);
        }

        private IExpressionBinding VisitDictionaryAccessExpression(Declaration module, Declaration parent, ParserRuleContext expression, ParserRuleContext nameContext, IExpressionBinding lExpressionBinding, ResolutionStatementContext statementContext)
        {
            /*
                A dictionary access expression is syntactically translated into an index expression with the same 
                expression for <l-expression> and an argument list with a single positional argument with a 
                declared type of String and a value equal to the name value of <unrestricted-name>. 
             */
            var fakeArgList = new ArgumentList();
            fakeArgList.AddArgument(new ArgumentListArgument(new LiteralDefaultBinding(nameContext), ArgumentListArgumentType.Positional));
            return new IndexDefaultBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression, lExpressionBinding, fakeArgList);
        }

        private IExpressionBinding VisitDictionaryAccessExpression(Declaration module, Declaration parent, ParserRuleContext expression, ParserRuleContext nameContext, IBoundExpression lExpression, ResolutionStatementContext statementContext)
        {
            /*
                A dictionary access expression is syntactically translated into an index expression with the same 
                expression for <l-expression> and an argument list with a single positional argument with a 
                declared type of String and a value equal to the name value of <unrestricted-name>. 
             */
            var fakeArgList = new ArgumentList();
            fakeArgList.AddArgument(new ArgumentListArgument(new LiteralDefaultBinding(nameContext), ArgumentListArgumentType.Positional));
            return new IndexDefaultBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression, lExpression, fakeArgList);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.WithExprContext expression, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            return Visit(module, parent, expression.withExpression(), withBlockVariable, ResolutionStatementContext.Undefined);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.WithExpressionContext expression, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            /*
                A <with-member-access-expression> or <with-dictionary-access-expression> is 
                statically resolved as a normal member access or dictionary access expression, respectively, as if 
                the innermost enclosing With block variable was specified for <l-expression>. If there is no 
                enclosing With block, the <with-expression> is invalid.
             */
            if (expression.withMemberAccessExpression() != null)
            {
                return new MemberAccessDefaultBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression,  withBlockVariable, expression.withMemberAccessExpression().unrestrictedName().GetText(), statementContext, expression.withMemberAccessExpression().unrestrictedName());
            }
            else
            {
                return VisitDictionaryAccessExpression(module, parent, expression, expression.withDictionaryAccessExpression().unrestrictedName(), withBlockVariable, ResolutionStatementContext.Undefined);
            }
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.ParenthesizedExprContext expression, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            dynamic expressionParens = expression.expression();
            var expressionBinding = Visit(module, parent, expressionParens, withBlockVariable, ResolutionStatementContext.Undefined);
            return new ParenthesizedDefaultBinding(expression, expressionBinding);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.ParenthesizedExpressionContext expression, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            dynamic expressionParens = expression.expression();
            var expressionBinding = Visit(module, parent, expressionParens, withBlockVariable, ResolutionStatementContext.Undefined);
            return new ParenthesizedDefaultBinding(expression, expressionBinding);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.TypeOfIsExprContext expression, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            return Visit(module, parent, expression.typeOfIsExpression(), withBlockVariable, ResolutionStatementContext.Undefined);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.TypeOfIsExpressionContext expression, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            dynamic booleanExpression = expression.expression();
            var booleanExpressionBinding = Visit(module, parent, booleanExpression, withBlockVariable, ResolutionStatementContext.Undefined);
            dynamic typeExpression = expression.typeExpression();
            var typeExpressionBinding = Visit(module, parent, typeExpression, withBlockVariable, ResolutionStatementContext.Undefined);
            return new TypeOfIsDefaultBinding(expression, booleanExpressionBinding, typeExpressionBinding);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.PowOpContext expression, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            return VisitBinaryOp(module, parent, expression, expression.expression()[0], expression.expression()[1], withBlockVariable, ResolutionStatementContext.Undefined);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.MultOpContext expression, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            return VisitBinaryOp(module, parent, expression, expression.expression()[0], expression.expression()[1], withBlockVariable, ResolutionStatementContext.Undefined);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.IntDivOpContext expression, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            return VisitBinaryOp(module, parent, expression, expression.expression()[0], expression.expression()[1], withBlockVariable, ResolutionStatementContext.Undefined);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.ModOpContext expression, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            return VisitBinaryOp(module, parent, expression, expression.expression()[0], expression.expression()[1], withBlockVariable, ResolutionStatementContext.Undefined);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.AddOpContext expression, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            return VisitBinaryOp(module, parent, expression, expression.expression()[0], expression.expression()[1], withBlockVariable, ResolutionStatementContext.Undefined);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.ConcatOpContext expression, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            return VisitBinaryOp(module, parent, expression, expression.expression()[0], expression.expression()[1], withBlockVariable, ResolutionStatementContext.Undefined);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.RelationalOpContext expression, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            return VisitBinaryOp(module, parent, expression, expression.expression()[0], expression.expression()[1], withBlockVariable, ResolutionStatementContext.Undefined);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.LogicalAndOpContext expression, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            return VisitBinaryOp(module, parent, expression, expression.expression()[0], expression.expression()[1], withBlockVariable, ResolutionStatementContext.Undefined);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.LogicalOrOpContext expression, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            return VisitBinaryOp(module, parent, expression, expression.expression()[0], expression.expression()[1], withBlockVariable, ResolutionStatementContext.Undefined);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.LogicalXorOpContext expression, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            return VisitBinaryOp(module, parent, expression, expression.expression()[0], expression.expression()[1], withBlockVariable, ResolutionStatementContext.Undefined);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.LogicalEqvOpContext expression, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            return VisitBinaryOp(module, parent, expression, expression.expression()[0], expression.expression()[1], withBlockVariable, ResolutionStatementContext.Undefined);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.LogicalImpOpContext expression, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            return VisitBinaryOp(module, parent, expression, expression.expression()[0], expression.expression()[1], withBlockVariable, ResolutionStatementContext.Undefined);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.UnaryMinusOpContext expression, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            return VisitUnaryOp(module, parent, expression, expression.expression(), withBlockVariable, ResolutionStatementContext.Undefined);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.LogicalNotOpContext expression, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            return VisitUnaryOp(module, parent, expression, expression.expression(), withBlockVariable, ResolutionStatementContext.Undefined);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.LiteralExprContext expression, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            return Visit(module, parent, expression.literalExpression(), withBlockVariable, ResolutionStatementContext.Undefined);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.InstanceExprContext expression, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            return Visit(module, parent, expression.instanceExpression(), withBlockVariable, ResolutionStatementContext.Undefined);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.InstanceExpressionContext expression, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            return new InstanceDefaultBinding(expression, module);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.LiteralExpressionContext expression, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            return new LiteralDefaultBinding(expression);
        }

        private IExpressionBinding VisitBinaryOp(Declaration module, Declaration parent, ParserRuleContext context, ParserRuleContext left, ParserRuleContext right, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            dynamic leftExpr = left;
            var leftBinding = Visit(module, parent, leftExpr, withBlockVariable, ResolutionStatementContext.Undefined);
            dynamic rightExpr = right;
            var rightBinding = Visit(module, parent, rightExpr, withBlockVariable, ResolutionStatementContext.Undefined);
            return new BinaryOpDefaultBinding(context, leftBinding, rightBinding);
        }

        private IExpressionBinding VisitUnaryOp(Declaration module, Declaration parent, ParserRuleContext context, ParserRuleContext expr, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            dynamic exprExpr = expr;
            var exprBinding = Visit(module, parent, exprExpr, withBlockVariable, ResolutionStatementContext.Undefined);
            return new UnaryOpDefaultBinding(context, exprBinding);
        }
    }
}
