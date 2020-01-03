using System;
using System.Collections.Generic;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Binding;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA.ReferenceManagement
{
    public sealed class BoundExpressionVisitor
    {
        private readonly DeclarationFinder _declarationFinder;

        public BoundExpressionVisitor(DeclarationFinder declarationFinder)
        {
            _declarationFinder = declarationFinder;
        }

        public void AddIdentifierReferences(
            IBoundExpression boundExpression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent,
            bool isAssignmentTarget = false,
            bool hasExplicitLetStatement = false,
            bool isSetAssignment = false)
        {
            Visit(boundExpression, module, scope, parent, isAssignmentTarget, hasExplicitLetStatement, isSetAssignment);
        }

        private void Visit(
            IBoundExpression boundExpression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent,
            bool isAssignmentTarget = false,
            bool hasExplicitLetStatement = false,
            bool isSetAssignment = false,
            bool hasArrayAccess = false)
        {
            switch (boundExpression)
            {
                case SimpleNameExpression simpleNameExpression:
                    Visit(simpleNameExpression, module, scope, parent, isAssignmentTarget, hasExplicitLetStatement, isSetAssignment);
                    break;
                case MemberAccessExpression memberAccessExpression:
                    Visit(memberAccessExpression, module, scope, parent, isAssignmentTarget, hasExplicitLetStatement, isSetAssignment);
                    break;
                case IndexExpression indexExpression:
                    Visit(indexExpression, module, scope, parent, isAssignmentTarget, hasExplicitLetStatement, isSetAssignment);
                    break;
                case ParenthesizedExpression parenthesizedExpression:
                    Visit(parenthesizedExpression, module, scope, parent);
                    break;
                case LiteralExpression literalExpression:
                    break;
                case BinaryOpExpression binaryOpExpression:
                    Visit(binaryOpExpression, module, scope, parent);
                    break;
                case UnaryOpExpression unaryOpExpression:
                    Visit(unaryOpExpression, module, scope, parent);
                    break;
                case NewExpression newExpression:
                    Visit(newExpression, module, scope, parent);
                    break;
                case InstanceExpression instanceExpression:
                    Visit(instanceExpression, module, scope, parent, isAssignmentTarget, hasExplicitLetStatement, isSetAssignment);
                    break;
                case DictionaryAccessExpression dictionaryAccessExpression:
                    Visit(dictionaryAccessExpression, module, scope, parent, isAssignmentTarget, hasExplicitLetStatement, isSetAssignment);
                    break;
                case TypeOfIsExpression typeOfIsExpression:
                    Visit(typeOfIsExpression, module, scope, parent);
                    break;
                case ResolutionFailedExpression resolutionFailedExpression:
                    Visit(resolutionFailedExpression, module, scope, parent);
                    break;
                case BuiltInTypeExpression builtInTypeExpression:
                    break;
                case RecursiveDefaultMemberAccessExpression recursiveDefaultMemberAccessExpression:
                    Visit(recursiveDefaultMemberAccessExpression, module, scope, parent, hasExplicitLetStatement, hasArrayAccess);
                    break;
                case LetCoercionDefaultMemberAccessExpression letCoercionDefaultMemberAccessExpression:
                    Visit(letCoercionDefaultMemberAccessExpression, module, scope, parent, isAssignmentTarget, hasExplicitLetStatement);
                    break;
                case ProcedureCoercionExpression procedureCoercionExpression:
                    Visit(procedureCoercionExpression, module, scope, parent);
                    break;
                case MissingArgumentExpression missingArgumentExpression:
                    break;
                case OutputListExpression outputListExpression:
                    Visit(outputListExpression, module, scope, parent);
                    break;
                case ObjectPrintExpression objectPrintExpression:
                    Visit(objectPrintExpression, module, scope, parent);
                    break;
                default:
                    throw new NotSupportedException($"Unexpected bound expression type {boundExpression.GetType()}");
            }
        }

        private void Visit(
            ResolutionFailedExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent)
        {
            // To bind as much as possible we gather all successfully resolved expressions and bind them here as a special case.
            foreach (var successfullyResolvedExpression in expression.SuccessfullyResolvedExpressions)
            {
                Visit(successfullyResolvedExpression, module, scope, parent);
            }
        }

        private void Visit(
            SimpleNameExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent,
            bool isAssignmentTarget,
            bool hasExplicitLetStatement,
            bool isSetAssignment)
        {
            var callSiteContext = expression.Context;
            var callee = expression.ReferencedDeclaration;
            var identifier = WithEnclosingBracketsRemoved(callSiteContext.GetText());
            var selection = callSiteContext.GetSelection();
            expression.ReferencedDeclaration.AddReference(
                module,
                scope,
                parent,
                callSiteContext,
                identifier,
                callee,
                selection,
                FindIdentifierAnnotations(module, selection.StartLine),
                isAssignmentTarget,
                hasExplicitLetStatement,
                isSetAssignment);
        }

        private IEnumerable<IParseTreeAnnotation> FindIdentifierAnnotations(QualifiedModuleName module, int line)
        {
            return _declarationFinder.FindAnnotations(module, line, AnnotationTarget.Identifier);
        }

        private void Visit(
            MemberAccessExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent,
            bool isAssignmentTarget,
            bool hasExplicitLetStatement,
            bool isSetAssignment)
        {
            Visit(expression.LExpression, module, scope, parent);
            
            // Expressions could be unbound thus not have a referenced declaration. The lexpression might still be bindable though.
            if (expression.Classification == ExpressionClassification.Unbound)
            {
                return;
            }

            var callSiteContext = expression.UnrestrictedNameContext;
            var callee = expression.ReferencedDeclaration;
            var identifier = WithEnclosingBracketsRemoved(callSiteContext.GetText());
            var selection = callSiteContext.GetSelection();
            expression.ReferencedDeclaration.AddReference(
                module,
                scope,
                parent,
                callSiteContext,
                identifier,
                callee,
                selection,
                FindIdentifierAnnotations(module, selection.StartLine),
                isAssignmentTarget,
                hasExplicitLetStatement,
                isSetAssignment);
        }

        private static string WithEnclosingBracketsRemoved(string identifierName)
        {
            if (identifierName.StartsWith("[") && identifierName.EndsWith("]"))
            {
                return identifierName.Substring(1, identifierName.Length - 2);
            }

            return identifierName;
        }

        private void Visit(
            ObjectPrintExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent)
        {
            var outputListExpression = expression.OutputListExpression;
            if (outputListExpression != null)
            {
                Visit(expression.OutputListExpression, module, scope, parent);
            }

            Visit(expression.PrintMethodExpressions, module, scope, parent);
        }

        private void Visit(
            OutputListExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent)
        {
            foreach (var itemExpression in expression.ItemExpressions)
            {
                Visit(itemExpression, module, scope, parent);
            }
        }

        private void Visit(
            IndexExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent,
            bool isAssignmentTarget,
            bool hasExplicitLetStatement,
            bool isSetAssignment)
        {
            var containedExpression = expression.ContainedDefaultMemberRecursionExpression;
            if (containedExpression != null)
            {
                Visit(containedExpression, module, scope, parent, hasExplicitLetStatement: hasExplicitLetStatement, hasArrayAccess: expression.IsArrayAccess);
            }

            if (expression.IsDefaultMemberAccess)
            {
                Visit(expression.LExpression, module, scope, parent, hasExplicitLetStatement: hasExplicitLetStatement);

                switch (expression.Classification)
                {
                    case ExpressionClassification.ResolutionFailed:
                        AddFailedIndexedDefaultMemberReference(expression, module, scope, parent, isAssignmentTarget, hasExplicitLetStatement, expression.ArgumentList.HasArguments);
                        break;
                    case ExpressionClassification.Unbound:
                        AddUnboundDefaultMemberReference(expression, module, scope, parent, isAssignmentTarget, hasExplicitLetStatement, isSetAssignment);
                        break;
                    default:
                        if (expression.ReferencedDeclaration != null)
                        {
                            AddDefaultMemberReference(expression, module, scope, parent, isAssignmentTarget, hasExplicitLetStatement, isSetAssignment);
                        }

                        break;
                }
            }
            else if (expression.Classification != ExpressionClassification.Unbound
                && expression.IsArrayAccess
                && expression.ReferencedDeclaration != null)
            {
                Visit(expression.LExpression, module, scope, parent, hasExplicitLetStatement: hasExplicitLetStatement);
                AddArrayAccessReference(expression, module, scope, parent, isAssignmentTarget, hasExplicitLetStatement, isSetAssignment);
            }
            else if (expression.Classification == ExpressionClassification.Unbound
                     && expression.ReferencedDeclaration == null)
            {
                Visit(expression.LExpression, module, scope, parent);
            }
            else
            {
                // Index expressions are a bit special in that they can refer to parameterized properties and functions.
                // In that case, the reference goes to the property or function. So, we pass on the assignment flags.
                Visit(expression.LExpression, module, scope, parent, isAssignmentTarget, hasExplicitLetStatement, isSetAssignment);
            }

            // Argument lists are not affected by the resolution of the target of the index expression.
            foreach (var argument in expression.ArgumentList.Arguments)
            {
                if (argument.ReferencedParameter != null)
                {
                    AddArgumentReference(argument, module, scope, parent);
                }

                if (argument.Expression != null)
                {
                    Visit(argument.Expression, module, scope, parent);
                }
                if (argument.NamedArgumentExpression != null)
                {
                    Visit(argument.NamedArgumentExpression, module, scope, parent);
                }
            }
        }

        private void AddArgumentReference(
            ArgumentListArgument argument,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent
        )
        {
            var argumentContext = argument.Context;
            var identifier = argumentContext.GetText();
            var argumentSelection = argumentContext.GetSelection();
            argument.ReferencedParameter.AddArgumentReference(
                module,
                scope,
                parent,
                argumentSelection,
                argumentContext,
                argument.ArgumentListContext,
                argument.ArgumentType,
                argument.ArgumentPosition,
                identifier,
                FindIdentifierAnnotations(module, argumentSelection.StartLine));
        }

        private void AddArrayAccessReference(
            IndexExpression expression, 
            QualifiedModuleName module, 
            Declaration scope,
            Declaration parent, 
            bool isAssignmentTarget, 
            bool hasExplicitLetStatement, 
            bool isSetAssignment)
        {
            var callSiteContext = expression.Context;
            var identifier = callSiteContext.GetText();
            var selection = callSiteContext.GetSelection();
            var callee = expression.ReferencedDeclaration;
            expression.ReferencedDeclaration.AddReference(
                module,
                scope,
                parent,
                callSiteContext,
                identifier,
                callee,
                selection,
                FindIdentifierAnnotations(module, selection.StartLine),
                isAssignmentTarget,
                hasExplicitLetStatement,
                isSetAssignment,
                isArrayAccess: true);
        }

        private void AddDefaultMemberReference(
            IndexExpression expression, 
            QualifiedModuleName module, 
            Declaration scope,
            Declaration parent, 
            bool isAssignmentTarget,
            bool hasExplicitLetStatement,
            bool isSetAssignment)
        {
            var callSiteContext = expression.LExpression.Context;
            var identifier = expression.Context.GetText();
            var selection = callSiteContext.GetSelection();
            var callee = expression.ReferencedDeclaration;
            expression.ReferencedDeclaration.AddReference(
                module,
                scope,
                parent,
                callSiteContext,
                identifier,
                callee,
                selection,
                FindIdentifierAnnotations(module, selection.StartLine),
                isAssignmentTarget,
                hasExplicitLetStatement,
                isSetAssignment,
                isIndexedDefaultMemberAccess: true,
                defaultMemberRecursionDepth: expression.DefaultMemberRecursionDepth);
        }

        private void AddUnboundDefaultMemberReference(
            IndexExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent,
            bool isAssignmentTarget,
            bool hasExplicitLetStatement,
            bool isSetAssignment)
        {
            var callSiteContext = expression.LExpression.Context;
            var identifier = expression.Context.GetText();
            var selection = callSiteContext.GetSelection();
            var callee = expression.ReferencedDeclaration;
            var reference = new IdentifierReference(
                module,
                scope,
                parent,
                identifier,
                selection,
                callSiteContext,
                callee,
                isAssignmentTarget,
                hasExplicitLetStatement,
                FindIdentifierAnnotations(module, selection.StartLine),
                isSetAssignment,
                isIndexedDefaultMemberAccess: true,
                defaultMemberRecursionDepth: expression.DefaultMemberRecursionDepth);
            _declarationFinder.AddUnboundDefaultMemberAccess(reference);
        }

        private void AddFailedIndexedDefaultMemberReference(
            IndexExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent,
            bool isAssignmentTarget,
            bool hasExplicitLetStatement,
            bool hasArguments)
        {
            var callSiteContext = expression.Context;
            var identifier = callSiteContext.GetText();
            var selection = callSiteContext.GetSelection();
            var callee = expression.ReferencedDeclaration;
            var reference = new IdentifierReference(
                module,
                scope,
                parent,
                identifier,
                selection,
                callSiteContext,
                callee,
                isAssignmentTarget,
                hasExplicitLetStatement,
                FindIdentifierAnnotations(module, selection.StartLine),
                false,
                isIndexedDefaultMemberAccess: hasArguments,
                isNonIndexedDefaultMemberAccess: !hasArguments,
                defaultMemberRecursionDepth: expression.DefaultMemberRecursionDepth);
            _declarationFinder.AddFailedIndexedDefaultMemberResolution(reference);
        }

        private void Visit(
            DictionaryAccessExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent,
            bool isAssignmentTarget,
            bool hasExplicitLetStatement,
            bool isSetAssignment)
        {
            Visit(expression.LExpression, module, scope, parent, hasExplicitLetStatement: hasExplicitLetStatement);

            var containedExpression = expression.ContainedDefaultMemberRecursionExpression;
            if (containedExpression != null)
            {
                Visit(containedExpression, module, scope, parent, hasExplicitLetStatement);
            }

            switch (expression.Classification)
            {
                case ExpressionClassification.ResolutionFailed:
                    AddFailedIndexedDefaultMemberReference(expression, module, scope, parent, isAssignmentTarget, hasExplicitLetStatement);
                    break;
                case ExpressionClassification.Unbound:
                    AddUnboundDefaultMemberReference(expression, module, scope, parent, isAssignmentTarget, hasExplicitLetStatement, isSetAssignment);
                    break;
                default:
                    if (expression.ReferencedDeclaration != null)
                    {
                        AddDefaultMemberReference(expression, module, scope, parent, isAssignmentTarget, hasExplicitLetStatement, isSetAssignment);
                    }

                    break;
            }

            // Argument List not affected by being unbound.
            foreach (var argument in expression.ArgumentList.Arguments)
            {
                if (argument.ReferencedParameter != null)
                {
                    AddArgumentReference(argument, module, scope, parent);
                }

                if (argument.Expression != null)
                {
                    Visit(argument.Expression, module, scope, parent);
                }
                if (argument.NamedArgumentExpression != null)
                {
                    Visit(argument.NamedArgumentExpression, module, scope, parent);
                }
            }
        }

        private void AddDefaultMemberReference(
            ProcedureCoercionExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent)
        {
            var callSiteContext = expression.Context;
            var identifier = callSiteContext.GetText();
            var selection = callSiteContext.GetSelection();
            var callee = expression.ReferencedDeclaration;
            expression.ReferencedDeclaration.AddReference(
                module,
                scope,
                parent,
                callSiteContext,
                identifier,
                callee,
                selection,
                FindIdentifierAnnotations(module, selection.StartLine),
                isNonIndexedDefaultMemberAccess: true,
                defaultMemberRecursionDepth: expression.DefaultMemberRecursionDepth,
                isProcedureCoercion: true);
        }

        private void AddUnboundDefaultMemberReference(
            ProcedureCoercionExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent)
        {
            var callSiteContext = expression.Context;
            var identifier = callSiteContext.GetText();
            var selection = callSiteContext.GetSelection();
            var callee = expression.ReferencedDeclaration;
            var reference = new IdentifierReference(
                module,
                scope,
                parent,
                identifier,
                selection,
                callSiteContext,
                callee,
                false,
                false,
                FindIdentifierAnnotations(module, selection.StartLine),
                false,
                isNonIndexedDefaultMemberAccess: true,
                defaultMemberRecursionDepth: expression.DefaultMemberRecursionDepth,
                isProcedureCoercion: true);
            _declarationFinder.AddUnboundDefaultMemberAccess(reference);
        }

        private void Visit(
            RecursiveDefaultMemberAccessExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent,
            bool hasExplicitLetStatement,
            bool hasArrayAccess)
        {
            var containedExpression = expression.ContainedDefaultMemberRecursionExpression;
            if (containedExpression != null)
            {
                Visit(containedExpression, module, scope, parent, hasExplicitLetStatement: hasExplicitLetStatement);
            }

            if (expression.Classification != ExpressionClassification.Unbound
                && expression.ReferencedDeclaration != null)
            {
                AddDefaultMemberReference(expression, module, parent, scope, hasExplicitLetStatement, !hasArrayAccess);
            }
        }

        private void AddDefaultMemberReference(
            RecursiveDefaultMemberAccessExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent,
            bool hasExplicitLetStatement,
            bool isInnerRecursiveDefaultMemberAccess)
        {
            var callSiteContext = expression.Context;
            var identifier = callSiteContext.GetText();
            var selection = callSiteContext.GetSelection();
            var callee = expression.ReferencedDeclaration;
            expression.ReferencedDeclaration.AddReference(
                module,
                scope,
                parent,
                callSiteContext,
                identifier,
                callee,
                selection,
                FindIdentifierAnnotations(module, selection.StartLine),
                hasExplicitLetStatement: hasExplicitLetStatement,
                isNonIndexedDefaultMemberAccess: true,
                defaultMemberRecursionDepth: expression.DefaultMemberRecursionDepth,
                isInnerRecursiveDefaultMemberAccess: isInnerRecursiveDefaultMemberAccess);
        }

        private void AddFailedIndexedDefaultMemberReference(
            DictionaryAccessExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent,
            bool isAssignmentTarget,
            bool hasExplicitLetStatement)
        {
            var callSiteContext = expression.Context;
            var identifier = callSiteContext.GetText();
            var selection = callSiteContext.GetSelection();
            var callee = expression.ReferencedDeclaration;
            var reference = new IdentifierReference(
                module,
                scope,
                parent,
                identifier,
                selection,
                callSiteContext,
                callee,
                isAssignmentTarget,
                hasExplicitLetStatement,
                FindIdentifierAnnotations(module, selection.StartLine),
                false,
                isIndexedDefaultMemberAccess: true,
                defaultMemberRecursionDepth: expression.DefaultMemberRecursionDepth);
            _declarationFinder.AddFailedIndexedDefaultMemberResolution(reference);
        }

        private void Visit(
            LetCoercionDefaultMemberAccessExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent,
            bool isAssignmentTarget = false,
            bool hasExplicitLetStatement = false)
        {
            var containedExpression = expression.ContainedDefaultMemberRecursionExpression;
            if (containedExpression != null)
            {
                Visit(containedExpression, module, scope, parent, hasExplicitLetStatement: hasExplicitLetStatement);
            }

            Visit(expression.WrappedExpression, module, scope, parent);

            switch (expression.Classification)
            {
                case ExpressionClassification.ResolutionFailed:
                    AddFailedLetCoercionReference(expression, module, scope, parent, isAssignmentTarget, hasExplicitLetStatement);
                    break;
                case ExpressionClassification.Unbound:
                    AddUnboundDefaultMemberReference(expression, module, scope, parent, isAssignmentTarget, hasExplicitLetStatement);
                    break;
                default:
                    if (expression.ReferencedDeclaration != null)
                    {
                        AddDefaultMemberReference(expression, module, scope, parent, isAssignmentTarget, hasExplicitLetStatement);
                    }

                    break;
            }
        }

        private void AddDefaultMemberReference(
            LetCoercionDefaultMemberAccessExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent,
            bool isAssignmentTarget,
            bool hasExplicitLetStatement)
        {
            var callSiteContext = expression.Context;
            var identifier = callSiteContext.GetText();
            var selection = callSiteContext.GetSelection();
            var callee = expression.ReferencedDeclaration;
            expression.ReferencedDeclaration.AddReference(
                module,
                scope,
                parent,
                callSiteContext,
                identifier,
                callee,
                selection,
                FindIdentifierAnnotations(module, selection.StartLine),
                isAssignmentTarget,
                hasExplicitLetStatement,
                isNonIndexedDefaultMemberAccess: true,
                defaultMemberRecursionDepth: expression.DefaultMemberRecursionDepth);
        }

        private void AddUnboundDefaultMemberReference(
            LetCoercionDefaultMemberAccessExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent,
            bool isAssignmentTarget,
            bool hasExplicitLetStatement)
        {
            var callSiteContext = expression.Context;
            var identifier = callSiteContext.GetText();
            var selection = callSiteContext.GetSelection();
            var callee = expression.ReferencedDeclaration;
            var reference = new IdentifierReference(
                module,
                scope,
                parent,
                identifier,
                selection,
                callSiteContext,
                callee,
                isAssignmentTarget,
                hasExplicitLetStatement,
                FindIdentifierAnnotations(module, selection.StartLine),
                false,
                isNonIndexedDefaultMemberAccess: true,
                defaultMemberRecursionDepth: expression.DefaultMemberRecursionDepth);
            _declarationFinder.AddUnboundDefaultMemberAccess(reference);
        }

        private void AddFailedLetCoercionReference(
            LetCoercionDefaultMemberAccessExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent,
            bool isAssignmentTarget,
            bool hasExplicitLetStatement)
        {
            var callSiteContext = expression.Context;
            var identifier = callSiteContext.GetText();
            var selection = callSiteContext.GetSelection();
            var callee = expression.ReferencedDeclaration;
            var reference = new IdentifierReference(
                module,
                scope,
                parent,
                identifier,
                selection,
                callSiteContext,
                callee,
                isAssignmentTarget,
                hasExplicitLetStatement,
                FindIdentifierAnnotations(module, selection.StartLine),
                false,
                isNonIndexedDefaultMemberAccess: true,
                defaultMemberRecursionDepth: expression.DefaultMemberRecursionDepth);
            _declarationFinder.AddFailedLetCoercionReference(reference);
        }

        private void Visit(
            ProcedureCoercionExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent)
        {
            Visit(expression.WrappedExpression, module, scope, parent);

            switch (expression.Classification)
            {
                case ExpressionClassification.ResolutionFailed:
                    AddFailedProcedureCoercionReference(expression, module, scope, parent);
                    break;
                case ExpressionClassification.Unbound:
                    AddUnboundDefaultMemberReference(expression, module, scope, parent);
                    break;
                default:
                    if (expression.ReferencedDeclaration != null)
                    {
                        AddDefaultMemberReference(expression, module, scope, parent);
                    }

                    break;
            }
        }

        private void AddDefaultMemberReference(
            DictionaryAccessExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent,
            bool isAssignmentTarget,
            bool hasExplicitLetStatement,
            bool isSetAssignment)
        {
            var callSiteContext = expression.DefaultMemberContext;
            var identifier = expression.Context.GetText();
            var callee = expression.ReferencedDeclaration;
            var selection = callSiteContext.GetSelection();
            expression.ReferencedDeclaration.AddReference(
                module,
                scope,
                parent,
                callSiteContext,
                identifier,
                callee,
                selection,
                FindIdentifierAnnotations(module, selection.StartLine),
                isAssignmentTarget,
                hasExplicitLetStatement,
                isSetAssignment,
                isIndexedDefaultMemberAccess: true,
                defaultMemberRecursionDepth: expression.DefaultMemberRecursionDepth);
        }

        private void AddUnboundDefaultMemberReference(
            DictionaryAccessExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent,
            bool isAssignmentTarget,
            bool hasExplicitLetStatement,
            bool isSetAssignment)
        {
            var callSiteContext = expression.DefaultMemberContext;
            var identifier = expression.Context.GetText();
            var selection = callSiteContext.GetSelection();
            var callee = expression.ReferencedDeclaration;
            var reference = new IdentifierReference(
                module,
                scope,
                parent,
                identifier,
                selection,
                callSiteContext,
                callee,
                isAssignmentTarget,
                hasExplicitLetStatement,
                FindIdentifierAnnotations(module, selection.StartLine),
                isSetAssignment,
                isIndexedDefaultMemberAccess: true,
                defaultMemberRecursionDepth: expression.DefaultMemberRecursionDepth);
            _declarationFinder.AddUnboundDefaultMemberAccess(reference);
        }

        private void AddFailedProcedureCoercionReference(
            ProcedureCoercionExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent)
        {
            var callSiteContext = expression.Context;
            var identifier = callSiteContext.GetText();
            var selection = callSiteContext.GetSelection();
            var callee = expression.ReferencedDeclaration;
            var reference = new IdentifierReference(
                module,
                scope,
                parent,
                identifier,
                selection,
                callSiteContext,
                callee,
                false,
                false,
                FindIdentifierAnnotations(module, selection.StartLine),
                false,
                isNonIndexedDefaultMemberAccess: true,
                defaultMemberRecursionDepth: expression.DefaultMemberRecursionDepth);
            _declarationFinder.AddFailedProcedureCoercionReference(reference);
        }

        private void Visit(
            NewExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent)
        {
            // We don't need to add a reference to the NewExpression's referenced declaration since that's covered
            // with its TypeExpression.
            Visit(expression.TypeExpression, module, scope, parent);
        }

        private void Visit(
            ParenthesizedExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent)
        {
            Visit(expression.Expression, module, scope, parent);
        }

        private void Visit(
            TypeOfIsExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent)
        {
            Visit(expression.Expression, module, scope, parent);
            Visit(expression.TypeExpression, module, scope, parent);
        }

        private void Visit(
            BinaryOpExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent)
        {
            Visit(expression.Left, module, scope, parent);
            Visit(expression.Right, module, scope, parent);
        }

        private void Visit(
            UnaryOpExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent)
        {
            Visit(expression.Expr, module, scope, parent);
        }

        private void Visit(
            InstanceExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent,
            bool isAssignmentTarget,
            bool hasExplicitLetStatement,
            bool isSetAssignment)
        {
            var callSiteContext = expression.Context;
            var identifier = expression.Context.GetText();
            var selection = callSiteContext.GetSelection();
            var callee = expression.ReferencedDeclaration;
            expression.ReferencedDeclaration.AddReference(
                module,
                scope,
                parent,
                callSiteContext,
                identifier,
                callee,
                selection,
                FindIdentifierAnnotations(module, selection.StartLine),
                isAssignmentTarget,
                hasExplicitLetStatement,
                isSetAssignment);
        }
    }
}
