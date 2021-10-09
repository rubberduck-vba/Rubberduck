using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using Rubberduck.Parsing.VBA.DeclarationCaching;

namespace Rubberduck.Parsing.Binding
{
    public sealed class DefaultBindingContext : BindingContextBase
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

        public override IBoundExpression Resolve(Declaration module,
            Declaration parent,
            ParserRuleContext expression,
            IBoundExpression withBlockVariable,
            StatementResolutionContext statementContext,
            bool requiresLetCoercion = false,
            bool isLetAssignment = false)
        {
            var bindingTree = BuildTree(module, parent, expression, withBlockVariable, statementContext, requiresLetCoercion, isLetAssignment);
            return bindingTree?.Resolve();
        }

        public override IExpressionBinding BuildTree(Declaration module,
            Declaration parent,
            ParserRuleContext expression,
            IBoundExpression withBlockVariable,
            StatementResolutionContext statementContext,
            bool requiresLetCoercion = false,
            bool isLetAssignment = false)
        {
            return Visit(
                module,
                parent,
                expression,
                withBlockVariable,
                statementContext,
                requiresLetCoercion,
                isLetAssignment);
        }

        public IExpressionBinding Visit(Declaration module, Declaration parent, ParserRuleContext expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext, bool requiresLetCoercion = false, bool isLetAssignment = false)
        {
            if (requiresLetCoercion && expression is ParserRuleContext context)
            {
                var innerExpressionBinding = BuildTree(module, parent, expression, withBlockVariable, statementContext);
                return new LetCoercionDefaultBinding(context, innerExpressionBinding, isLetAssignment);
            }

            switch (expression)
            {
                case VBAParser.ExpressionContext expressionContext:
                    return Visit(module, parent, expressionContext, withBlockVariable, statementContext);
                case VBAParser.LExpressionContext lExpressionContext:
                    return Visit(module, parent, lExpressionContext, withBlockVariable, statementContext);
                case VBAParser.IdentifierValueContext identifierValueContext:
                    return Visit(module, parent, identifierValueContext, statementContext);
                case VBAParser.CallStmtContext callExpression:
                    return Visit(module, parent, callExpression, withBlockVariable);
                case VBAParser.BooleanExpressionContext booleanExpressionContext:
                    return Visit(module, parent, booleanExpressionContext, withBlockVariable);
                case VBAParser.IntegerExpressionContext integerExpressionContext:
                    return Visit(module, parent, integerExpressionContext, withBlockVariable);
                case VBAParser.OutputListContext outputListContext:
                    return Visit(module, parent, outputListContext, withBlockVariable);
                case VBAParser.UnqualifiedObjectPrintStmtContext unqualifiedObjectPrintStmtContext:
                    return Visit(module, parent, unqualifiedObjectPrintStmtContext, withBlockVariable);
                default:
                    return HandleUnexpectedExpressionType(expression);
            }
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.CallStmtContext expression, IBoundExpression withBlockVariable)
        {
            var lExpression = expression.lExpression();
            var lExpressionBinding = Visit(module, parent, lExpression, withBlockVariable, StatementResolutionContext.Undefined);

            if (expression.CALL() == null)
            {
                var argList = VisitArgumentList(module, parent, expression.argumentList(), withBlockVariable);
                SetLeftMatch(lExpressionBinding, argList.Arguments.Count);
                if (argList.HasArguments)
                {
                    return new IndexDefaultBinding(expression.lExpression(), lExpressionBinding, argList, parent);
                }

                return new ProcedureCoercionDefaultBinding(expression.lExpression(), lExpressionBinding, false, parent);
            }

            return new ProcedureCoercionDefaultBinding(expression.lExpression(), lExpressionBinding, true, parent);
        }

        private static void SetLeftMatch(IExpressionBinding binding, int argumentCount)
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

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.ExpressionContext expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            switch (expression)
            {
                case VBAParser.LExprContext lExprContext:
                    return Visit(module, parent, lExprContext.lExpression(), withBlockVariable, statementContext);
                case VBAParser.ParenthesizedExprContext parenthesizedExprContext:
                    return Visit(module, parent, parenthesizedExprContext, withBlockVariable);
                case VBAParser.RelationalOpContext relationalOpContext:
                    return Visit(module, parent, relationalOpContext, withBlockVariable);
                case VBAParser.LiteralExprContext literalExprContext:
                    return Visit(literalExprContext.literalExpression());
                case VBAParser.NewExprContext newExprContext:
                    return Visit(module, parent, newExprContext, withBlockVariable);
                case VBAParser.LogicalNotOpContext logicalNotOpContext:
                    return VisitUnaryOp(module, parent, logicalNotOpContext, logicalNotOpContext.expression(), withBlockVariable);
                case VBAParser.UnaryMinusOpContext unaryMinusOpContext:
                    return VisitUnaryOp(module, parent, unaryMinusOpContext, unaryMinusOpContext.expression(), withBlockVariable);
                case VBAParser.LogicalAndOpContext logicalAndOpContext:
                    return VisitBinaryOp(module, parent, logicalAndOpContext, logicalAndOpContext.expression()[0], logicalAndOpContext.expression()[1], withBlockVariable);
                case VBAParser.LogicalOrOpContext logicalOrOpContext:
                    return VisitBinaryOp(module, parent, logicalOrOpContext, logicalOrOpContext.expression()[0], logicalOrOpContext.expression()[1], withBlockVariable);
                case VBAParser.LogicalXorOpContext logicalXorOpContext:
                    return VisitBinaryOp(module, parent, logicalXorOpContext, logicalXorOpContext.expression()[0], logicalXorOpContext.expression()[1], withBlockVariable);
                case VBAParser.LogicalEqvOpContext logicalEqvOpContext:
                    return VisitBinaryOp(module, parent, logicalEqvOpContext, logicalEqvOpContext.expression()[0], logicalEqvOpContext.expression()[1], withBlockVariable);
                case VBAParser.LogicalImpOpContext logicalImpOpContext:
                    return VisitBinaryOp(module, parent, logicalImpOpContext, logicalImpOpContext.expression()[0], logicalImpOpContext.expression()[1], withBlockVariable);
                case VBAParser.AddOpContext addOpContext:
                    return VisitBinaryOp(module, parent, addOpContext, addOpContext.expression()[0], addOpContext.expression()[1], withBlockVariable);
                case VBAParser.ConcatOpContext concatOpContext:
                    return VisitBinaryOp(module, parent, concatOpContext, concatOpContext.expression()[0], concatOpContext.expression()[1], withBlockVariable);
                case VBAParser.MultOpContext multOpContext:
                    return VisitBinaryOp(module, parent, multOpContext, multOpContext.expression()[0], multOpContext.expression()[1], withBlockVariable);
                case VBAParser.ModOpContext modOpContext:
                    return VisitBinaryOp(module, parent, modOpContext, modOpContext.expression()[0], modOpContext.expression()[1], withBlockVariable);
                case VBAParser.PowOpContext powOpContext:
                    return VisitBinaryOp(module, parent, powOpContext, powOpContext.expression()[0], powOpContext.expression()[1], withBlockVariable);
                case VBAParser.IntDivOpContext intDivOpContext:
                    return VisitBinaryOp(module, parent, intDivOpContext, intDivOpContext.expression()[0], intDivOpContext.expression()[1], withBlockVariable);
                case VBAParser.MarkedFileNumberExprContext markedFileNumberExprContext:
                    return Visit(module, parent, markedFileNumberExprContext, withBlockVariable);
                case VBAParser.BuiltInTypeExprContext builtInTypeExprContext:
                    return Visit(module, parent, builtInTypeExprContext, statementContext);
                //We do not handle the VBAParser.TypeofexprContext because that should only ever appear as a child of an IS relational operator expression and is specifically handled there.
                default:
                    return HandleUnexpectedExpressionType(expression);
            }
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.LExpressionContext expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            switch (expression)
            {
                case VBAParser.SimpleNameExprContext simpleNameExprContext:
                    return Visit(module, parent, simpleNameExprContext, statementContext);
                case VBAParser.MemberAccessExprContext memberAccessExprContext:
                    return Visit(module, parent, memberAccessExprContext, withBlockVariable, statementContext);
                case VBAParser.IndexExprContext indexExprContext:
                    return Visit(module, parent, indexExprContext, withBlockVariable);
                case VBAParser.WithMemberAccessExprContext withMemberAccessExprContext:
                    return Visit(module, parent, withMemberAccessExprContext, withBlockVariable, statementContext);
                case VBAParser.InstanceExprContext instanceExprContext:
                    return Visit(module, instanceExprContext);
                case VBAParser.WhitespaceIndexExprContext whitespaceIndexExprContext:
                    return Visit(module, parent, whitespaceIndexExprContext, withBlockVariable);
                case VBAParser.DictionaryAccessExprContext dictionaryAccessExprContext:
                    return Visit(module, parent, dictionaryAccessExprContext, withBlockVariable);
                case VBAParser.WithDictionaryAccessExprContext withDictionaryAccessExprContext:
                    return Visit(module, parent, withDictionaryAccessExprContext, withBlockVariable);
                case VBAParser.ObjectPrintExprContext objectPrintExprContext:
                    return Visit(module, parent, objectPrintExprContext, withBlockVariable);
                default:
                    return HandleUnexpectedExpressionType(expression);
            }
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.NewExprContext expression, IBoundExpression withBlockVariable)
        {
            var typeExpressionBinding = VisitType(module, parent, expression.expression(), withBlockVariable);
            if (typeExpressionBinding == null)
            {
                return null;
            }
            return new NewTypeBinding(expression, typeExpressionBinding);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.MarkedFileNumberExprContext expression, IBoundExpression withBlockVariable)
        {
            // The MarkedFileNumberExpr doesn't actually exist but for backwards compatibility reasons we support it, ignore the "hash tag" of the file number
            // and resolve it as a normal expression.
            // This allows us to support functions such as Input(file1, #file1) which would otherwise not work.
            return Visit(module, parent, expression.expression(), withBlockVariable, StatementResolutionContext.Undefined);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.BuiltInTypeExprContext expression, StatementResolutionContext statementContext)
        {
            //It is legal to name variables like built-in types.
            //So, we try and see whether it resolves.
            var builtInTypeNamedNameBinding = new SimpleNameDefaultBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression, Identifier.GetName(expression.builtInType()), statementContext);
            var resolutionResult = builtInTypeNamedNameBinding.Resolve();
            if (resolutionResult.Classification != ExpressionClassification.ResolutionFailed)
            {
                return builtInTypeNamedNameBinding;
            }

            // Not actually an expression, but treated as one to allow for a faster parser.
            return new BuiltInTypeDefaultBinding(expression);
        }

        private IExpressionBinding VisitType(Declaration module, Declaration parent, VBAParser.ExpressionContext expression, IBoundExpression withBlockVariable)
        {
            return _typeBindingContext.BuildTree(module, parent, expression, withBlockVariable, StatementResolutionContext.Undefined);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.SimpleNameExprContext expression, StatementResolutionContext statementContext)
        {
            return new SimpleNameDefaultBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression, Identifier.GetName(expression.identifier()), statementContext);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.IdentifierValueContext expression, StatementResolutionContext statementContext)
        {
            return new SimpleNameDefaultBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression, Identifier.GetName(expression), statementContext);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.MemberAccessExprContext expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            var lExpression = expression.lExpression();
            var lExpressionBinding = Visit(module, parent, lExpression, withBlockVariable, StatementResolutionContext.Undefined);
            return new MemberAccessDefaultBinding(
                _declarationFinder, 
                Declaration.GetProjectParent(parent), 
                module, 
                parent, 
                expression, 
                lExpressionBinding,
                statementContext, 
                expression.unrestrictedIdentifier());
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.ObjectPrintExprContext expression, IBoundExpression withBlockVariable)
        {
            var lExpression = expression.lExpression();
            var lExpressionBinding = Visit(module, parent, lExpression, withBlockVariable, StatementResolutionContext.Undefined);
            var memberAccessBinding = new MemberAccessDefaultBinding(
                _declarationFinder,
                Declaration.GetProjectParent(parent),
                module,
                parent,
                expression,
                lExpressionBinding,
                StatementResolutionContext.Undefined,
                expression.printMethod());
            var outputListContext = expression.outputList();
            var outputListBinding = outputListContext != null
                ? Visit(
                    module,
                    parent,
                    outputListContext,
                    withBlockVariable)
                : null;
            return new ObjectPrintDefaultBinding(expression, memberAccessBinding, outputListBinding);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.UnqualifiedObjectPrintStmtContext expression, IBoundExpression withBlockVariable)
        {
            var printMethodContext = expression.printMethod();
            var simpleNameBinding = new SimpleNameDefaultBinding(
                _declarationFinder,
                Declaration.GetProjectParent(parent),
                module,
                parent,
                printMethodContext,
                printMethodContext.GetText(),
                StatementResolutionContext.Undefined);
            var outputListContext = expression.outputList();
            var outputListBinding = outputListContext != null 
                ? Visit(
                    module,
                    parent,
                    outputListContext,
                    withBlockVariable)
                : null;
            return new ObjectPrintDefaultBinding(expression, simpleNameBinding, outputListBinding);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.IndexExprContext expression, IBoundExpression withBlockVariable)
        {
            var lExpression = expression.lExpression();
            var lExpressionBinding = Visit(module, parent, lExpression, withBlockVariable, StatementResolutionContext.Undefined);
            var argumentListBinding = VisitArgumentList(module, parent, expression.argumentList(), withBlockVariable);
            SetLeftMatch(lExpressionBinding, argumentListBinding.Arguments.Count);
            return new IndexDefaultBinding(expression, lExpressionBinding, argumentListBinding, parent);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.WhitespaceIndexExprContext expression, IBoundExpression withBlockVariable)
        {
            var lExpression = expression.lExpression();
            var lExpressionBinding = Visit(module, parent, lExpression, withBlockVariable, StatementResolutionContext.Undefined);
            var argumentListBinding = VisitArgumentList(module, parent, expression.argumentList(), withBlockVariable);
            SetLeftMatch(lExpressionBinding, argumentListBinding.Arguments.Count);
            return new IndexDefaultBinding(expression, lExpressionBinding, argumentListBinding, parent);
        }

        private ArgumentList VisitArgumentList(Declaration module, Declaration parent, VBAParser.ArgumentListContext argumentList, IBoundExpression withBlockVariable)
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
                        var (binding, context, isAddressOfArgument) = VisitArgumentBinding(module, parent, expr.positionalArgument().argumentExpression(), withBlockVariable);
                        convertedList.AddArgument(new ArgumentListArgument(
                            binding, 
                            context,
                            argumentList,
                            ArgumentListArgumentType.Positional, 
                            isAddressOfArgument));
                    }
                    else if (expr.namedArgument() != null)
                    {
                        var (binding, context, isAddressOfArgument) = VisitArgumentBinding(module, parent, expr.namedArgument().argumentExpression(), withBlockVariable);
                        convertedList.AddArgument(new ArgumentListArgument(
                            binding, 
                            context, 
                            argumentList,
                            ArgumentListArgumentType.Named,
                            CreateNamedArgumentExpressionCreator(expr.namedArgument().unrestrictedIdentifier().GetText(), expr.namedArgument().unrestrictedIdentifier()),
                            isAddressOfArgument));
                    }
                    else if(expr.missingArgument() != null)
                    {
                        var missingArgumentContext = expr.missingArgument();
                        var binding = new MissingArgumentBinding(missingArgumentContext);
                        convertedList.AddArgument(new ArgumentListArgument(
                            binding,
                            missingArgumentContext,
                            argumentList,
                            ArgumentListArgumentType.Missing,
                            false));
                    }
                }
            }
            return convertedList;
        }

        private Func<Declaration, IBoundExpression> CreateNamedArgumentExpressionCreator(string parameterName, ParserRuleContext context)
        {
            return calledProcedure =>
            {
                var classification = ExpressionClassificationOfProcedure(calledProcedure);
                var parameter = _declarationFinder.FindParameter(calledProcedure, parameterName);
                return parameter != null 
                    ? new SimpleNameExpression(parameter, classification, context) 
                    : null;
            };
        }

        private static ExpressionClassification ExpressionClassificationOfProcedure(Declaration procedure)
        {
            switch (procedure.DeclarationType)
            {
                case DeclarationType.Procedure:
                    return ExpressionClassification.Subroutine;
                case DeclarationType.Function:
                case DeclarationType.LibraryFunction:
                case DeclarationType.LibraryProcedure:
                    return ExpressionClassification.Function;
                default:
                    return ExpressionClassification.Property;
            }
        }

        private (IExpressionBinding binding, ParserRuleContext context, bool isAddressOfArgument) VisitArgumentBinding(Declaration module, Declaration parent, VBAParser.ArgumentExpressionContext argumentExpression, IBoundExpression withBlockVariable)
        {
            if (argumentExpression.expression() != null)
            {
                var expr = argumentExpression.expression();
                return (Visit(module, parent, expr, withBlockVariable, StatementResolutionContext.Undefined), expr, false);
            }
            else
            {
                var expr = argumentExpression.addressOfExpression();
                return (Visit(module, parent, expr, withBlockVariable, StatementResolutionContext.Undefined), expr, true);
            }
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.AddressOfExpressionContext expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            return _procedurePointerBindingContext.BuildTree(module, parent, expression, withBlockVariable, statementContext);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.DictionaryAccessExprContext expression, IBoundExpression withBlockVariable)
        {
            var lExpression = expression.lExpression();
            var lExpressionBinding = Visit(module, parent, lExpression, withBlockVariable, StatementResolutionContext.Undefined);
            return VisitDictionaryAccessExpression(expression, expression.unrestrictedIdentifier(), lExpressionBinding);
        }

        private IExpressionBinding VisitDictionaryAccessExpression(ParserRuleContext expression, ParserRuleContext nameContext, IExpressionBinding lExpressionBinding)
        {
            /*
                A dictionary access expression is syntactically translated into an index expression with the same 
                expression for <l-expression> and an argument list with a single positional argument with a 
                declared type of String and a value equal to the name value of <unrestricted-name>. 

                Still, we have a specific binding for it in order to attach a reference to the called default member to the exclamation mark.
             */
            var fakeArgList = new ArgumentList();
            fakeArgList.AddArgument(new ArgumentListArgument(new LiteralDefaultBinding(nameContext), nameContext, null, ArgumentListArgumentType.Positional));
            return new DictionaryAccessDefaultBinding(expression, lExpressionBinding, fakeArgList);
        }

        private IExpressionBinding VisitDictionaryAccessExpression(ParserRuleContext expression, ParserRuleContext nameContext, IBoundExpression lExpression)
        {
            /*
                A dictionary access expression is syntactically translated into an index expression with the same 
                expression for <l-expression> and an argument list with a single positional argument with a 
                declared type of String and a value equal to the name value of <unrestricted-name>. 

                Still, we have a specific binding for it in order to attach a reference to the called default member to the exclamation mark.
             */
            var fakeArgList = new ArgumentList();
            fakeArgList.AddArgument(new ArgumentListArgument(new LiteralDefaultBinding(nameContext), nameContext, null, ArgumentListArgumentType.Positional));
            return new DictionaryAccessDefaultBinding(expression, lExpression, fakeArgList);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.WithMemberAccessExprContext expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext)
        {
            if (withBlockVariable == null)
            {
                withBlockVariable = new ResolutionFailedExpression(expression);
            }
            return new MemberAccessDefaultBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression, withBlockVariable, expression.unrestrictedIdentifier().GetText(), statementContext, expression.unrestrictedIdentifier());
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.WithDictionaryAccessExprContext expression, IBoundExpression withBlockVariable)
        {
            /*
                A <with-member-access-expression> or <with-dictionary-access-expression> is 
                statically resolved as a normal member access or dictionary access expression, respectively, as if 
                the innermost enclosing With block variable was specified for <l-expression>. If there is no 
                enclosing With block, the <with-expression> is invalid.
             */
            return VisitDictionaryAccessExpression(expression, expression.unrestrictedIdentifier(), withBlockVariable);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.ParenthesizedExprContext expression, IBoundExpression withBlockVariable)
        {
            var expressionParens = expression.expression();
            var expressionBinding = Visit(module, parent, expressionParens, withBlockVariable, StatementResolutionContext.Undefined);
            var letCoercionBinding = new LetCoercionDefaultBinding(expressionParens, expressionBinding);
            return new ParenthesizedDefaultBinding(expression, letCoercionBinding);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.RelationalOpContext expression, IBoundExpression withBlockVariable)
        {
            // To make the grammar we treat a type-of-is expression as a construct of the form "TYPEOF expression", where expression
            // is always "expression IS expression".
            if (expression.expression()[0] is VBAParser.TypeofexprContext typeofExpr)
            {
                return VisitTypeOf(module, parent, expression, typeofExpr, expression.expression()[1], withBlockVariable);
            }

            var isIsRelation = expression.IS() != null;
            return VisitBinaryOp(module, parent, expression, expression.expression()[0], expression.expression()[1], withBlockVariable, isIsRelation);
        }

        private IExpressionBinding VisitTypeOf(
            Declaration module, 
            Declaration parent,
            VBAParser.RelationalOpContext typeOfIsExpression,
            VBAParser.TypeofexprContext typeOfLeftPartExpression,
            VBAParser.ExpressionContext typeExpression,
            IBoundExpression withBlockVariable)
        {
            var booleanExpression = typeOfLeftPartExpression.expression();
            var booleanExpressionBinding = Visit(module, parent, booleanExpression, withBlockVariable, StatementResolutionContext.Undefined);
            var typeExpressionBinding = VisitType(module, parent, typeExpression, withBlockVariable);
            return new TypeOfIsDefaultBinding(typeOfIsExpression, booleanExpressionBinding, typeExpressionBinding);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.BooleanExpressionContext expression, IBoundExpression withBlockVariable)
        {
            var innerExpression = Visit(module, parent, expression.expression(), withBlockVariable, StatementResolutionContext.Undefined);
            return new LetCoercionDefaultBinding(expression, innerExpression);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.IntegerExpressionContext expression, IBoundExpression withBlockVariable)
        {
            var innerExpression =  Visit(module, parent, expression.expression(), withBlockVariable, StatementResolutionContext.Undefined);
            return new LetCoercionDefaultBinding(expression, innerExpression);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAParser.OutputListContext outputListContext, IBoundExpression withBlockVariable)
        {
            var itemBindings = new List<IExpressionBinding>();
            foreach (var outputItem in outputListContext.outputItem())
            {
                if (outputItem.outputClause() != null)
                {
                    if (outputItem.outputClause().spcClause() != null)
                    {
                        itemBindings.Add(Visit(
                            module, 
                            parent, 
                            outputItem.outputClause().spcClause().spcNumber().expression(), 
                            withBlockVariable, 
                            StatementResolutionContext.Undefined, 
                            requiresLetCoercion: true));
                    }
                    if (outputItem.outputClause().tabClause() != null && outputItem.outputClause().tabClause().tabNumberClause() != null)
                    {
                        itemBindings.Add(Visit(
                            module, 
                            parent, 
                            outputItem.outputClause().tabClause().tabNumberClause().tabNumber().expression(), 
                            withBlockVariable, 
                            StatementResolutionContext.Undefined, 
                            requiresLetCoercion: true));
                    }
                    if (outputItem.outputClause().outputExpression() != null)
                    {
                        itemBindings.Add(Visit(
                            module,
                            parent,
                            outputItem.outputClause().outputExpression().expression(),
                            withBlockVariable,
                            StatementResolutionContext.Undefined,
                            requiresLetCoercion: true));
                    }
                }
            }

            return new OutputListDefaultBinding(outputListContext, itemBindings);
        }

        private static IExpressionBinding Visit(Declaration module, VBAParser.InstanceExprContext expression)
        {
            return new InstanceDefaultBinding(expression, module);
        }

        private static IExpressionBinding Visit(VBAParser.LiteralExpressionContext expression)
        {
            return new LiteralDefaultBinding(expression);
        }

        private IExpressionBinding VisitBinaryOp(Declaration module, Declaration parent, ParserRuleContext context, VBAParser.ExpressionContext left, VBAParser.ExpressionContext right, IBoundExpression withBlockVariable, bool isIsRelation = false)
        {
            var leftBinding = Visit(module, parent, left, withBlockVariable, StatementResolutionContext.Undefined);
            var rightBinding = Visit(module, parent, right, withBlockVariable, StatementResolutionContext.Undefined);

            if (!isIsRelation)
            {
                //All but the Is relation require value type arguments.
                leftBinding = new LetCoercionDefaultBinding(left, leftBinding);
                rightBinding = new LetCoercionDefaultBinding(right, rightBinding);
            }

            return new BinaryOpDefaultBinding(context, leftBinding, rightBinding);
        }

        private IExpressionBinding VisitUnaryOp(Declaration module, Declaration parent, ParserRuleContext context, VBAParser.ExpressionContext expr, IBoundExpression withBlockVariable)
        {
            var exprBinding = Visit(module, parent, expr, withBlockVariable, StatementResolutionContext.Undefined);
            var letCoercionBinding = new LetCoercionDefaultBinding(expr, exprBinding);
            return new UnaryOpDefaultBinding(context, letCoercionBinding);
        }
    }
}
