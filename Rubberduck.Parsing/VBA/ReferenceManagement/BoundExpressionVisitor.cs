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
            bool isSetAssignment = false,
            bool isReDim = false)
        {
            Visit(boundExpression, module, scope, parent, isAssignmentTarget, hasExplicitLetStatement, isSetAssignment, isReDim: isReDim);
        }

        /// <summary>
        /// Traverses the tree of expressions and adds identifier references to the referenced declarations.
        /// It returns the reference that another directly depended reference would relate to, e.g. if that reference is a member access. 
        /// </summary>
        /// <param name="boundExpression"></param>
        /// <param name="module"></param>
        /// <param name="scope"></param>
        /// <param name="parent"></param>
        /// <param name="isAssignmentTarget"></param>
        /// <param name="hasExplicitLetStatement"></param>
        /// <param name="isSetAssignment"></param>
        /// <param name="hasArrayAccess"></param>
        /// <returns>The reference in an expression a member or array access on it would refer to.</returns>
        private IdentifierReference Visit(
            IBoundExpression boundExpression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent,
            bool isAssignmentTarget = false,
            bool hasExplicitLetStatement = false,
            bool isSetAssignment = false,
            bool hasArrayAccess = false,
            bool isReDim = false)
        {
            switch (boundExpression)
            {
                case SimpleNameExpression simpleNameExpression:
                    return Visit(simpleNameExpression, module, scope, parent, isAssignmentTarget, hasExplicitLetStatement, isSetAssignment, isReDim);
                case MemberAccessExpression memberAccessExpression:
                    return Visit(memberAccessExpression, module, scope, parent, isAssignmentTarget, hasExplicitLetStatement, isSetAssignment);
                case IndexExpression indexExpression:
                    return Visit(indexExpression, module, scope, parent, isAssignmentTarget, hasExplicitLetStatement, isSetAssignment);
                case ParenthesizedExpression parenthesizedExpression:
                    return Visit(parenthesizedExpression, module, scope, parent);
                case LiteralExpression literalExpression:
                    return null;
                case BinaryOpExpression binaryOpExpression:
                    return Visit(binaryOpExpression, module, scope, parent);
                case UnaryOpExpression unaryOpExpression:
                    return Visit(unaryOpExpression, module, scope, parent);
                case NewExpression newExpression:
                    return Visit(newExpression, module, scope, parent);
                case InstanceExpression instanceExpression:
                    return Visit(instanceExpression, module, scope, parent, isAssignmentTarget, hasExplicitLetStatement, isSetAssignment);
                case DictionaryAccessExpression dictionaryAccessExpression:
                    return Visit(dictionaryAccessExpression, module, scope, parent, isAssignmentTarget, hasExplicitLetStatement, isSetAssignment);
                case TypeOfIsExpression typeOfIsExpression:
                    return Visit(typeOfIsExpression, module, scope, parent);
                case ResolutionFailedExpression resolutionFailedExpression:
                    return Visit(resolutionFailedExpression, module, scope, parent);
                case BuiltInTypeExpression builtInTypeExpression:
                    return null;
                case RecursiveDefaultMemberAccessExpression recursiveDefaultMemberAccessExpression:
                    return Visit(recursiveDefaultMemberAccessExpression, module, scope, parent, hasExplicitLetStatement, hasArrayAccess);
                case LetCoercionDefaultMemberAccessExpression letCoercionDefaultMemberAccessExpression:
                    return Visit(letCoercionDefaultMemberAccessExpression, module, scope, parent, isAssignmentTarget, hasExplicitLetStatement);
                case ProcedureCoercionExpression procedureCoercionExpression:
                    return Visit(procedureCoercionExpression, module, scope, parent);
                case MissingArgumentExpression missingArgumentExpression:
                    return null;
                case OutputListExpression outputListExpression:
                    return Visit(outputListExpression, module, scope, parent);
                case ObjectPrintExpression objectPrintExpression:
                    return Visit(objectPrintExpression, module, scope, parent);
                default:
                    throw new NotSupportedException($"Unexpected bound expression type {boundExpression.GetType()}");
            }
        }

        private IdentifierReference Visit(
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

            return null;
        }

        private IdentifierReference Visit(
            SimpleNameExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent,
            bool isAssignmentTarget,
            bool hasExplicitLetStatement,
            bool isSetAssignment,
            bool isReDim)
        {
            var callSiteContext = expression.Context;
            var callee = expression.ReferencedDeclaration;
            var identifier = WithEnclosingBracketsRemoved(callSiteContext.GetText());
            var selection = callSiteContext.GetSelection();
            return expression.ReferencedDeclaration.AddReference(
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
                isReDim: isReDim);
        }

        private IEnumerable<IParseTreeAnnotation> FindIdentifierAnnotations(QualifiedModuleName module, int line)
        {
            return _declarationFinder.FindAnnotations(module, line, AnnotationTarget.Identifier);
        }

        private IdentifierReference Visit(
            MemberAccessExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent,
            bool isAssignmentTarget,
            bool hasExplicitLetStatement,
            bool isSetAssignment)
        {
            var qualifyingReference = Visit(expression.LExpression, module, scope, parent);
            
            // Expressions could be unbound thus not have a referenced declaration. The lexpression might still be bindable though.
            if (expression.Classification == ExpressionClassification.Unbound)
            {
                return null;
            }

            var callSiteContext = expression.UnrestrictedNameContext;
            var callee = expression.ReferencedDeclaration;
            var identifier = WithEnclosingBracketsRemoved(callSiteContext.GetText());
            var selection = callSiteContext.GetSelection();
            return expression.ReferencedDeclaration.AddReference(
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
                qualifyingReference: qualifyingReference);
        }

        private static string WithEnclosingBracketsRemoved(string identifierName)
        {
            if (identifierName.StartsWith("[") && identifierName.EndsWith("]"))
            {
                return identifierName.Substring(1, identifierName.Length - 2);
            }

            return identifierName;
        }

        private IdentifierReference Visit(
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

            return null;
        }

        private IdentifierReference Visit(
            OutputListExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent)
        {
            foreach (var itemExpression in expression.ItemExpressions)
            {
                Visit(itemExpression, module, scope, parent);
            }

            return null;
        }

        private IdentifierReference Visit(
            IndexExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent,
            bool isAssignmentTarget,
            bool hasExplicitLetStatement,
            bool isSetAssignment)
        {
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

            if (expression.IsDefaultMemberAccess)
            {
                var qualifyingReference = Visit(expression.LExpression, module, scope, parent, hasExplicitLetStatement: hasExplicitLetStatement);

                var containedExpression = expression.ContainedDefaultMemberRecursionExpression;
                if (containedExpression != null)
                {
                    qualifyingReference = Visit(containedExpression, module, scope, parent, hasExplicitLetStatement: hasExplicitLetStatement, hasArrayAccess: expression.IsArrayAccess, initialQualifyingReference: qualifyingReference);
                }

                switch (expression.Classification)
                {
                    case ExpressionClassification.ResolutionFailed:
                        return AddFailedIndexedDefaultMemberReference(expression, module, scope, parent, isAssignmentTarget, hasExplicitLetStatement, expression.ArgumentList.HasArguments, qualifyingReference);
                    case ExpressionClassification.Unbound:
                        return AddUnboundDefaultMemberReference(expression, module, scope, parent, isAssignmentTarget, hasExplicitLetStatement, isSetAssignment, qualifyingReference);
                    default:
                        return expression.ReferencedDeclaration != null 
                            ? AddDefaultMemberReference(expression, module, scope, parent, isAssignmentTarget, hasExplicitLetStatement, isSetAssignment, qualifyingReference) 
                            : null;
                }
            }
            
            if (expression.Classification != ExpressionClassification.Unbound
                && expression.IsArrayAccess
                && expression.ReferencedDeclaration != null)
            {
                var qualifyingReference = Visit(expression.LExpression, module, scope, parent, hasExplicitLetStatement: hasExplicitLetStatement);
                
                //The final default member might not have parameters and return an array.
                var containedExpression = expression.ContainedDefaultMemberRecursionExpression;
                if (containedExpression != null)
                {
                    qualifyingReference = Visit(containedExpression, module, scope, parent, hasExplicitLetStatement: hasExplicitLetStatement, hasArrayAccess: expression.IsArrayAccess, initialQualifyingReference: qualifyingReference);
                }

                return AddArrayAccessReference(expression, module, scope, parent, isAssignmentTarget, hasExplicitLetStatement, isSetAssignment, qualifyingReference);
            }
            
            if (expression.Classification == ExpressionClassification.Unbound
                     && expression.ReferencedDeclaration == null)
            {
                Visit(expression.LExpression, module, scope, parent);
                return null;
            }

            // Index expressions are a bit special in that they can refer to parameterized properties and functions.
            // In that case, the reference goes to the property or function. So, we pass on the assignment flags.
            return Visit(expression.LExpression, module, scope, parent, isAssignmentTarget, hasExplicitLetStatement, isSetAssignment);
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

        private IdentifierReference AddArrayAccessReference(
            IndexExpression expression, 
            QualifiedModuleName module, 
            Declaration scope,
            Declaration parent, 
            bool isAssignmentTarget, 
            bool hasExplicitLetStatement, 
            bool isSetAssignment,
            IdentifierReference qualifyingReference)
        {
            var callSiteContext = expression.Context;
            var identifier = callSiteContext.GetText();
            var selection = callSiteContext.GetSelection();
            var callee = expression.ReferencedDeclaration;
            return expression.ReferencedDeclaration.AddReference(
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
                isArrayAccess: true,
                qualifyingReference: qualifyingReference);
        }

        private IdentifierReference AddDefaultMemberReference(
            IndexExpression expression, 
            QualifiedModuleName module, 
            Declaration scope,
            Declaration parent, 
            bool isAssignmentTarget,
            bool hasExplicitLetStatement,
            bool isSetAssignment,
            IdentifierReference qualifyingReference)
        {
            var callSiteContext = expression.LExpression.Context;
            var identifier = expression.Context.GetText();
            var selection = callSiteContext.GetSelection();
            var callee = expression.ReferencedDeclaration;
            return expression.ReferencedDeclaration.AddReference(
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
                defaultMemberRecursionDepth: expression.DefaultMemberRecursionDepth,
                qualifyingReference: qualifyingReference);
        }

        private IdentifierReference AddUnboundDefaultMemberReference(
            IndexExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent,
            bool isAssignmentTarget,
            bool hasExplicitLetStatement,
            bool isSetAssignment,
            IdentifierReference qualifyingReference)
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
                defaultMemberRecursionDepth: expression.DefaultMemberRecursionDepth,
                qualifyingReference: qualifyingReference);
            _declarationFinder.AddUnboundDefaultMemberAccess(reference);
            return reference;
        }

        private IdentifierReference AddFailedIndexedDefaultMemberReference(
            IndexExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent,
            bool isAssignmentTarget,
            bool hasExplicitLetStatement,
            bool hasArguments,
            IdentifierReference qualifyIdentifierReference)
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
                defaultMemberRecursionDepth: expression.DefaultMemberRecursionDepth,
                qualifyingReference: qualifyIdentifierReference);
            _declarationFinder.AddFailedIndexedDefaultMemberResolution(reference);
            return reference;
        }

        private IdentifierReference Visit(
            DictionaryAccessExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent,
            bool isAssignmentTarget,
            bool hasExplicitLetStatement,
            bool isSetAssignment)
        {
            // Argument List not affected by the resolution of the dictionary access.
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

            var qualifyingReference = Visit(expression.LExpression, module, scope, parent, hasExplicitLetStatement: hasExplicitLetStatement);

            var containedExpression = expression.ContainedDefaultMemberRecursionExpression;
            if (containedExpression != null)
            {
                qualifyingReference = Visit(containedExpression, module, scope, parent, hasExplicitLetStatement, false, qualifyingReference);
            }

            switch (expression.Classification)
            {
                case ExpressionClassification.ResolutionFailed:
                    return AddFailedIndexedDefaultMemberReference(expression, module, scope, parent, isAssignmentTarget, hasExplicitLetStatement, qualifyingReference);
                case ExpressionClassification.Unbound:
                    return AddUnboundDefaultMemberReference(expression, module, scope, parent, isAssignmentTarget, hasExplicitLetStatement, isSetAssignment);
                default:
                    return expression.ReferencedDeclaration != null 
                        ? AddDefaultMemberReference(expression, module, scope, parent, isAssignmentTarget, hasExplicitLetStatement, isSetAssignment, qualifyingReference) 
                        : null;
            }
        }

        private IdentifierReference AddDefaultMemberReference(
            ProcedureCoercionExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent,
            IdentifierReference qualifyingReference)
        {
            var callSiteContext = expression.Context;
            var identifier = callSiteContext.GetText();
            var selection = callSiteContext.GetSelection();
            var callee = expression.ReferencedDeclaration;
            return expression.ReferencedDeclaration.AddReference(
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
                isProcedureCoercion: true,
                qualifyingReference: qualifyingReference);
        }

        private IdentifierReference AddUnboundDefaultMemberReference(
            ProcedureCoercionExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent,
            IdentifierReference qualifyingReference)
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
                isProcedureCoercion: true,
                qualifyingReference: qualifyingReference);
            _declarationFinder.AddUnboundDefaultMemberAccess(reference);
            return reference;
        }

        private IdentifierReference Visit(
            RecursiveDefaultMemberAccessExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent,
            bool hasExplicitLetStatement,
            bool hasArrayAccess,
            IdentifierReference initialQualifyingReference)
        {
            var containedExpression = expression.ContainedDefaultMemberRecursionExpression;
            var qualifyingReference = containedExpression == null
                ? initialQualifyingReference
                : Visit(
                    containedExpression, 
                    module, 
                    scope, 
                    parent, 
                    hasExplicitLetStatement,
                    false, 
                    initialQualifyingReference);

            if (expression.Classification != ExpressionClassification.Unbound
                && expression.ReferencedDeclaration != null)
            {
                return AddDefaultMemberReference(expression, module, parent, scope, hasExplicitLetStatement, !hasArrayAccess, qualifyingReference);
            }

            return null;
        }

        private IdentifierReference AddDefaultMemberReference(
            RecursiveDefaultMemberAccessExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent,
            bool hasExplicitLetStatement,
            bool isInnerRecursiveDefaultMemberAccess,
            IdentifierReference qualifyingReference)
        {
            var callSiteContext = expression.Context;
            var identifier = callSiteContext.GetText();
            var selection = callSiteContext.GetSelection();
            var callee = expression.ReferencedDeclaration;
            return expression.ReferencedDeclaration.AddReference(
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
                isInnerRecursiveDefaultMemberAccess: isInnerRecursiveDefaultMemberAccess,
                qualifyingReference: qualifyingReference);
        }

        private IdentifierReference AddFailedIndexedDefaultMemberReference(
            DictionaryAccessExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent,
            bool isAssignmentTarget,
            bool hasExplicitLetStatement,
            IdentifierReference qualifyingReference)
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
                defaultMemberRecursionDepth: expression.DefaultMemberRecursionDepth,
                qualifyingReference: qualifyingReference);
            _declarationFinder.AddFailedIndexedDefaultMemberResolution(reference);
            return reference;
        }

        private IdentifierReference Visit(
            LetCoercionDefaultMemberAccessExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent,
            bool isAssignmentTarget = false,
            bool hasExplicitLetStatement = false)
        {
            var qualifyingReference = Visit(expression.WrappedExpression, module, scope, parent);

            var containedExpression = expression.ContainedDefaultMemberRecursionExpression;
            if (containedExpression != null)
            {
                qualifyingReference = Visit(containedExpression, module, scope, parent, hasExplicitLetStatement: hasExplicitLetStatement, false, qualifyingReference);
            }

            switch (expression.Classification)
            {
                case ExpressionClassification.ResolutionFailed:
                    return AddFailedLetCoercionReference(expression, module, scope, parent, isAssignmentTarget, hasExplicitLetStatement, qualifyingReference);
                case ExpressionClassification.Unbound:
                    return AddUnboundDefaultMemberReference(expression, module, scope, parent, isAssignmentTarget, hasExplicitLetStatement, qualifyingReference);
                default:
                    return expression.ReferencedDeclaration != null
                        ? AddDefaultMemberReference(expression, module, scope, parent, isAssignmentTarget, hasExplicitLetStatement, qualifyingReference)
                        : null;
            }
        }

        private IdentifierReference AddDefaultMemberReference(
            LetCoercionDefaultMemberAccessExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent,
            bool isAssignmentTarget,
            bool hasExplicitLetStatement,
            IdentifierReference qualifyingReference)
        {
            var callSiteContext = expression.Context;
            var identifier = callSiteContext.GetText();
            var selection = callSiteContext.GetSelection();
            var callee = expression.ReferencedDeclaration;
            return expression.ReferencedDeclaration.AddReference(
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
                defaultMemberRecursionDepth: expression.DefaultMemberRecursionDepth,
                qualifyingReference: qualifyingReference);
        }

        private IdentifierReference AddUnboundDefaultMemberReference(
            LetCoercionDefaultMemberAccessExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent,
            bool isAssignmentTarget,
            bool hasExplicitLetStatement,
            IdentifierReference qualifyingReference)
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
                defaultMemberRecursionDepth: expression.DefaultMemberRecursionDepth,
                qualifyingReference: qualifyingReference);
            _declarationFinder.AddUnboundDefaultMemberAccess(reference);
            return reference; 
        }

        private IdentifierReference AddFailedLetCoercionReference(
            LetCoercionDefaultMemberAccessExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent,
            bool isAssignmentTarget,
            bool hasExplicitLetStatement,
            IdentifierReference qualifyingReference)
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
                defaultMemberRecursionDepth: expression.DefaultMemberRecursionDepth,
                qualifyingReference: qualifyingReference);
            _declarationFinder.AddFailedLetCoercionReference(reference);
            return reference;
        }

        private IdentifierReference Visit(
            ProcedureCoercionExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent)
        {
            var qualifyingReference = Visit(expression.WrappedExpression, module, scope, parent);

            //Although a procedure reference should never qualify anything, it does not hurt to return the reference here.
            switch (expression.Classification)
            {
                case ExpressionClassification.ResolutionFailed:
                    return AddFailedProcedureCoercionReference(expression, module, scope, parent, qualifyingReference);
                case ExpressionClassification.Unbound:
                    return AddUnboundDefaultMemberReference(expression, module, scope, parent, qualifyingReference);
                default:
                    if (expression.ReferencedDeclaration != null)
                    {
                        return AddDefaultMemberReference(expression, module, scope, parent, qualifyingReference);
                    }

                    return null;
            }
        }

        private IdentifierReference AddDefaultMemberReference(
            DictionaryAccessExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent,
            bool isAssignmentTarget,
            bool hasExplicitLetStatement,
            bool isSetAssignment,
            IdentifierReference qualifyingReference)
        {
            var callSiteContext = expression.DefaultMemberContext;
            var identifier = expression.Context.GetText();
            var callee = expression.ReferencedDeclaration;
            var selection = callSiteContext.GetSelection();
            return expression.ReferencedDeclaration.AddReference(
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
                defaultMemberRecursionDepth: expression.DefaultMemberRecursionDepth,
                qualifyingReference: qualifyingReference);
        }

        private IdentifierReference AddUnboundDefaultMemberReference(
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
            return reference;
        }

        private IdentifierReference AddFailedProcedureCoercionReference(
            ProcedureCoercionExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent,
            IdentifierReference qualifyingReference)
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
                qualifyingReference: qualifyingReference);
            _declarationFinder.AddFailedProcedureCoercionReference(reference);
            return reference;
        }

        private IdentifierReference Visit(
            NewExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent)
        {
            // We don't need to add a reference to the NewExpression's referenced declaration since that's covered
            // with its TypeExpression.
            Visit(expression.TypeExpression, module, scope, parent);
            return null;
        }

        private IdentifierReference Visit(
            ParenthesizedExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent)
        {
            return Visit(expression.Expression, module, scope, parent);
        }

        private IdentifierReference Visit(
            TypeOfIsExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent)
        {
            Visit(expression.Expression, module, scope, parent);
            Visit(expression.TypeExpression, module, scope, parent);
            return null;
        }

        private IdentifierReference Visit(
            BinaryOpExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent)
        {
            Visit(expression.Left, module, scope, parent);
            Visit(expression.Right, module, scope, parent);
            return null;
        }

        private IdentifierReference Visit(
            UnaryOpExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent)
        {
            Visit(expression.Expr, module, scope, parent);
            return null;
        }

        private IdentifierReference Visit(
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
            return expression.ReferencedDeclaration.AddReference(
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
