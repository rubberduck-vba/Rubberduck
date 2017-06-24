using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
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
            return Visit(module, parent, dynamicExpression, withBlockVariable, statementContext);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.CallStmtContext expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            // Call statements always have an argument list.
            // One of the reasons we're doing this is that an empty argument list could represent a call to a default member,
            // which requires us to use an IndexDefaultBinding.
            if (expression.CALL() == null)
            {
                dynamic lexpr = expression.lExpression();
                var lexprBinding = Visit(module, parent, lexpr, withBlockVariable, StatementResolutionContext.Undefined);
                var argList = VisitArgumentList(module, parent, expression.argumentList(), withBlockVariable, StatementResolutionContext.Undefined);
                SetLeftMatch(lexprBinding, argList.Arguments.Count);
                return new IndexDefaultBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression.lExpression(), lexprBinding, argList);
            }
            else
            {
                var lexprBinding = Visit(module, parent, (dynamic)expression.lExpression(), withBlockVariable, StatementResolutionContext.Undefined);
                if (!(lexprBinding is IndexDefaultBinding))
                {
                    return new IndexDefaultBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression.lExpression(), lexprBinding, new ArgumentList());
                }
                else
                {
                    return lexprBinding;
                }
            }
        }

        private void SetLeftMatch(IExpressionBinding binding, int argumentCount)
        {
            // See SimpleNameDefaultBinding for a description on why we're doing this.
            if (!(binding is SimpleNameDefaultBinding))
            {
                return;
            }
            if (argumentCount != 2)
            {
                return;
            }
            ((SimpleNameDefaultBinding)binding).IsPotentialLeftMatch = true;
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.LExprContext expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            dynamic lexpr = expression.lExpression();
            return Visit(module, parent, lexpr, withBlockVariable, statementContext);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.NewExprContext expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            var typeExpressionBinding = VisitType(module, parent, expression.expression(), withBlockVariable, StatementResolutionContext.Undefined);
            if (typeExpressionBinding == null)
            {
                return null;
            }
            return new NewTypeBinding(_declarationFinder, module, parent, expression, typeExpressionBinding);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.MarkedFileNumberExprContext expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            // The MarkedFileNumberExpr doesn't actually exist but for backwards compatibility reasons we support it, ignore the "hash tag" of the file number
            // and resolve it as a normal expression.
            // This allows us to support functions such as Input(file1, #file1) which would otherwise not work.
            return Visit(module, parent, (dynamic)expression.expression(), withBlockVariable, StatementResolutionContext.Undefined);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.BuiltInTypeExprContext expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            // Not actually an expression, but treated as one to allow for a faster parser.
            return null;
        }

        private IExpressionBinding VisitType(Declaration module, Declaration parent, ParserRuleContext expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            return _typeBindingContext.BuildTree(module, parent, expression, withBlockVariable, StatementResolutionContext.Undefined);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.SimpleNameExprContext expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            return new SimpleNameDefaultBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression, Identifier.GetName(expression.identifier()), statementContext);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.IdentifierValueContext expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            return new SimpleNameDefaultBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression, Identifier.GetName(expression), statementContext);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.MemberAccessExprContext expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            dynamic lExpression = expression.lExpression();
            var lExpressionBinding = Visit(module, parent, lExpression, withBlockVariable, StatementResolutionContext.Undefined);
            return new MemberAccessDefaultBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression, lExpressionBinding, statementContext, expression.unrestrictedIdentifier());
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.IndexExprContext expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            dynamic lExpression = expression.lExpression();
            var lExpressionBinding = Visit(module, parent, lExpression, withBlockVariable, StatementResolutionContext.Undefined);
            var argumentListBinding = VisitArgumentList(module, parent, expression.argumentList(), withBlockVariable, StatementResolutionContext.Undefined);
            SetLeftMatch(lExpressionBinding, argumentListBinding.Arguments.Count);
            return new IndexDefaultBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression, lExpressionBinding, argumentListBinding);
        }

        private ArgumentList VisitArgumentList(Declaration module, Declaration parent, VBAParser.ArgumentListContext argumentList, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            var convertedList = new ArgumentList();
            if (argumentList == null)
            {
                return convertedList;
            }
            var list = argumentList;
            // TODO: positionalArgumentOrMissing is there as preparation for argument compatibility checking.
            if (list.argument() != null)
            {
                foreach (var expr in list.argument())
                {
                    if (expr.positionalArgument() != null)
                    {
                        convertedList.AddArgument(new ArgumentListArgument(
                            VisitArgumentBinding(module, parent, expr.positionalArgument().argumentExpression(), withBlockVariable,
                            StatementResolutionContext.Undefined), ArgumentListArgumentType.Positional));
                    }
                    else if (expr.namedArgument() != null)
                    {
                        convertedList.AddArgument(new ArgumentListArgument(
                            VisitArgumentBinding(module, parent, expr.namedArgument().argumentExpression(), withBlockVariable,
                            StatementResolutionContext.Undefined), ArgumentListArgumentType.Named,
                            CreateNamedArgumentExpressionCreator(expr.namedArgument().unrestrictedIdentifier().GetText(), expr.namedArgument().unrestrictedIdentifier())));
                    }
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

        private IExpressionBinding VisitArgumentBinding(Declaration module, Declaration parent, VBAParser.ArgumentExpressionContext argumentExpression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            if (argumentExpression.expression() != null)
            {
                dynamic expr = argumentExpression.expression();
                return Visit(module, parent, expr, withBlockVariable, StatementResolutionContext.Undefined);
            }
            else
            {
                dynamic expr = argumentExpression.addressOfExpression();
                return Visit(module, parent, expr, withBlockVariable, StatementResolutionContext.Undefined);
            }
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.AddressOfExpressionContext expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            return _procedurePointerBindingContext.BuildTree(module, parent, expression, withBlockVariable, statementContext);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.DictionaryAccessExprContext expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            dynamic lExpression = expression.lExpression();
            var lExpressionBinding = Visit(module, parent, lExpression, withBlockVariable, StatementResolutionContext.Undefined);
            return VisitDictionaryAccessExpression(module, parent, expression, expression.unrestrictedIdentifier(), lExpressionBinding, StatementResolutionContext.Undefined);
        }

        private IExpressionBinding VisitDictionaryAccessExpression(Declaration module, Declaration parent, ParserRuleContext expression, ParserRuleContext nameContext, IExpressionBinding lExpressionBinding, StatementResolutionContext statementContext)
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

        private IExpressionBinding VisitDictionaryAccessExpression(Declaration module, Declaration parent, ParserRuleContext expression, ParserRuleContext nameContext, IBoundExpression lExpression, StatementResolutionContext statementContext)
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

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.WithMemberAccessExprContext expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            return new MemberAccessDefaultBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression, withBlockVariable, expression.unrestrictedIdentifier().GetText(), statementContext, expression.unrestrictedIdentifier());
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.WithDictionaryAccessExprContext expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            /*
                A <with-member-access-expression> or <with-dictionary-access-expression> is 
                statically resolved as a normal member access or dictionary access expression, respectively, as if 
                the innermost enclosing With block variable was specified for <l-expression>. If there is no 
                enclosing With block, the <with-expression> is invalid.
             */
            return VisitDictionaryAccessExpression(module, parent, expression, expression.unrestrictedIdentifier(), withBlockVariable, StatementResolutionContext.Undefined);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.ParenthesizedExprContext expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            dynamic expressionParens = expression.expression();
            var expressionBinding = Visit(module, parent, expressionParens, withBlockVariable, StatementResolutionContext.Undefined);
            return new ParenthesizedDefaultBinding(expression, expressionBinding);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.PowOpContext expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            return VisitBinaryOp(module, parent, expression, expression.expression()[0], expression.expression()[1], withBlockVariable, StatementResolutionContext.Undefined);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.MultOpContext expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            return VisitBinaryOp(module, parent, expression, expression.expression()[0], expression.expression()[1], withBlockVariable, StatementResolutionContext.Undefined);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.IntDivOpContext expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            return VisitBinaryOp(module, parent, expression, expression.expression()[0], expression.expression()[1], withBlockVariable, StatementResolutionContext.Undefined);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.ModOpContext expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            return VisitBinaryOp(module, parent, expression, expression.expression()[0], expression.expression()[1], withBlockVariable, StatementResolutionContext.Undefined);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.AddOpContext expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            return VisitBinaryOp(module, parent, expression, expression.expression()[0], expression.expression()[1], withBlockVariable, StatementResolutionContext.Undefined);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.ConcatOpContext expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            return VisitBinaryOp(module, parent, expression, expression.expression()[0], expression.expression()[1], withBlockVariable, StatementResolutionContext.Undefined);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.RelationalOpContext expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            // To make the grammar we treat a type-of-is expression as a construct of the form "TYPEOF expression", where expression
            // is always "expression IS expression".
            if (expression.expression()[0] is VBAParser.TypeofexprContext)
            {
                return VisitTypeOf(module, parent, expression, (VBAParser.TypeofexprContext)expression.expression()[0], expression.expression()[1], withBlockVariable, StatementResolutionContext.Undefined);
            }
            return VisitBinaryOp(module, parent, expression, expression.expression()[0], expression.expression()[1], withBlockVariable, StatementResolutionContext.Undefined);
        }

        private IExpressionBinding VisitTypeOf(
            Declaration module, 
            Declaration parent,
            VBAParser.RelationalOpContext typeOfIsExpression,
            VBAParser.TypeofexprContext typeOfLeftPartExpression,
            ParserRuleContext typeExpression,
            IBoundExpression withBlockVariable, 
            StatementResolutionContext statementContext)
        {
            dynamic booleanExpression = typeOfLeftPartExpression.expression();
            var booleanExpressionBinding = Visit(module, parent, booleanExpression, withBlockVariable, StatementResolutionContext.Undefined);
            var typeExpressionBinding = VisitType(module, parent, (dynamic)typeExpression, withBlockVariable, StatementResolutionContext.Undefined);
            return new TypeOfIsDefaultBinding(typeOfIsExpression, booleanExpressionBinding, typeExpressionBinding);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.LogicalAndOpContext expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            return VisitBinaryOp(module, parent, expression, expression.expression()[0], expression.expression()[1], withBlockVariable, StatementResolutionContext.Undefined);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.LogicalOrOpContext expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            return VisitBinaryOp(module, parent, expression, expression.expression()[0], expression.expression()[1], withBlockVariable, StatementResolutionContext.Undefined);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.LogicalXorOpContext expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            return VisitBinaryOp(module, parent, expression, expression.expression()[0], expression.expression()[1], withBlockVariable, StatementResolutionContext.Undefined);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.LogicalEqvOpContext expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            return VisitBinaryOp(module, parent, expression, expression.expression()[0], expression.expression()[1], withBlockVariable, StatementResolutionContext.Undefined);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.LogicalImpOpContext expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            return VisitBinaryOp(module, parent, expression, expression.expression()[0], expression.expression()[1], withBlockVariable, StatementResolutionContext.Undefined);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.UnaryMinusOpContext expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            return VisitUnaryOp(module, parent, expression, expression.expression(), withBlockVariable, StatementResolutionContext.Undefined);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.LogicalNotOpContext expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            return VisitUnaryOp(module, parent, expression, expression.expression(), withBlockVariable, StatementResolutionContext.Undefined);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.LiteralExprContext expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            return Visit(module, parent, expression.literalExpression(), withBlockVariable, StatementResolutionContext.Undefined);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.BooleanExpressionContext expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            return Visit(module, parent, (dynamic)expression.expression(), withBlockVariable, StatementResolutionContext.Undefined);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.IntegerExpressionContext expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            return Visit(module, parent, (dynamic)expression.expression(), withBlockVariable, StatementResolutionContext.Undefined);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.InstanceExprContext expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            return new InstanceDefaultBinding(expression, module);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.LiteralExpressionContext expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            return new LiteralDefaultBinding(expression);
        }

        private IExpressionBinding VisitBinaryOp(Declaration module, Declaration parent, ParserRuleContext context, ParserRuleContext left, ParserRuleContext right, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            dynamic leftExpr = left;
            var leftBinding = Visit(module, parent, leftExpr, withBlockVariable, StatementResolutionContext.Undefined);
            dynamic rightExpr = right;
            var rightBinding = Visit(module, parent, rightExpr, withBlockVariable, StatementResolutionContext.Undefined);
            return new BinaryOpDefaultBinding(context, leftBinding, rightBinding);
        }

        private IExpressionBinding VisitUnaryOp(Declaration module, Declaration parent, ParserRuleContext context, ParserRuleContext expr, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            dynamic exprExpr = expr;
            var exprBinding = Visit(module, parent, exprExpr, withBlockVariable, StatementResolutionContext.Undefined);
            return new UnaryOpDefaultBinding(context, exprBinding);
        }
    }
}
