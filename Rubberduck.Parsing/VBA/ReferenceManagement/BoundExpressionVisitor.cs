using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Binding;
using Rubberduck.Parsing.Grammar;
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
            bool isSetAssignment = false)
        {
            switch (boundExpression)
            {
                case SimpleNameExpression simpleNameExpression:
                    Visit(simpleNameExpression, module, scope, parent, isAssignmentTarget, hasExplicitLetStatement, isSetAssignment);
                    break;
                case MemberAccessExpression memberAccessExpression:
                    Visit(memberAccessExpression, module, scope, parent, isAssignmentTarget, hasExplicitLetStatement, isSetAssignment);
                    break;
                case IndexExpression failedExpression:
                    Visit(failedExpression, module, scope, parent, isAssignmentTarget, hasExplicitLetStatement, isSetAssignment);
                    break;
                case ParenthesizedExpression parenthesizedExpression:
                    Visit(parenthesizedExpression, module, scope, parent);
                    break;
                case LiteralExpression literalExpression:
                    Visit(literalExpression);
                    break;
                case BinaryOpExpression binaryOpExpression:
                    Visit(binaryOpExpression, module, scope, parent);
                    break;
                case UnaryOpExpression unaryOpExpression:
                    Visit(unaryOpExpression, module, scope, parent);
                    break;
                case NewExpression failedExpression:
                    Visit(failedExpression, module, scope, parent);
                    break;
                case InstanceExpression instanceExpression:
                    Visit(instanceExpression, module, scope, parent, isAssignmentTarget, hasExplicitLetStatement, isSetAssignment);
                    break;
                case TypeOfIsExpression typeOfIsExpression:
                    Visit(typeOfIsExpression, module, scope, parent);
                    break;
                case ResolutionFailedExpression resolutionFailedExpression:
                    Visit(resolutionFailedExpression, module, scope, parent);
                    break;
                default: throw new NotSupportedException($"Unexpected bound expression type {boundExpression.GetType()}");
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
            if (isAssignmentTarget && expression.Context.Parent is VBAParser.IndexExprContext && !expression.ReferencedDeclaration.IsArray)
            {
                // 'SomeDictionary' is not the assignment target in 'SomeDictionary("key") = 42'
                // ..but we want to treat array index assignment as assignment to the array itself.
                isAssignmentTarget = false;
                isSetAssignment = false;
            }

            var callSiteContext = expression.Context;
            var identifier = expression.Context.GetText();
            var callee = expression.ReferencedDeclaration;
            expression.ReferencedDeclaration.AddReference(
                module,
                scope,
                parent,
                callSiteContext,
                identifier,
                callee,
                callSiteContext.GetSelection(),
                FindIdentifierAnnotations(module, callSiteContext.GetSelection().StartLine),
                isAssignmentTarget,
                hasExplicitLetStatement,
                isSetAssignment);
        }

        private IEnumerable<IAnnotation> FindIdentifierAnnotations(QualifiedModuleName module, int line)
        {
            return _declarationFinder.FindAnnotations(module, line)
                .Where(annotation => annotation.AnnotationType.HasFlag(AnnotationType.IdentifierAnnotation));
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
            var identifier = expression.UnrestrictedNameContext.GetText();
            var callee = expression.ReferencedDeclaration;
            expression.ReferencedDeclaration.AddReference(
                module,
                scope,
                parent,
                callSiteContext,
                identifier,
                callee,
                callSiteContext.GetSelection(),
                FindIdentifierAnnotations(module, callSiteContext.GetSelection().StartLine),
                isAssignmentTarget,
                hasExplicitLetStatement,
                isSetAssignment);
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
            // Index expressions are a bit special in that they could refer to elements of an array, what apparently we don't want to
            // add an identifier reference to, that's why we pass on the isassignment/hasexplicitletstatement values.
            Visit(expression.LExpression, module, scope, parent, isAssignmentTarget, hasExplicitLetStatement, isSetAssignment);

            if (expression.Classification != ExpressionClassification.Unbound
                && expression.ReferencedDeclaration != null
                && !ReferenceEquals(expression.LExpression.ReferencedDeclaration, expression.ReferencedDeclaration))
            {
                // Referenced declaration could also be null if e.g. it's an array and the array is a "base type" such as String.
                if (expression.ReferencedDeclaration != null)
                {
                    var callSiteContext = expression.LExpression.Context;
                    var identifier = expression.LExpression.Context.GetText();
                    var callee = expression.ReferencedDeclaration;
                    expression.ReferencedDeclaration.AddReference(
                        module,
                        scope,
                        parent,
                        callSiteContext,
                        identifier,
                        callee,
                        callSiteContext.GetSelection(),
                        FindIdentifierAnnotations(module, callSiteContext.GetSelection().StartLine),
                        isSetAssignment);
                }
            }
            // Argument List not affected by being unbound.
            foreach (var argument in expression.ArgumentList.Arguments)
            {
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

        private void Visit(LiteralExpression expression)
        {
            // Nothing to do here.
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
            var callee = expression.ReferencedDeclaration;
            expression.ReferencedDeclaration.AddReference(
                module,
                scope,
                parent,
                callSiteContext,
                identifier,
                callee,
                callSiteContext.GetSelection(),
                FindIdentifierAnnotations(module, callSiteContext.GetSelection().StartLine),
                isAssignmentTarget,
                hasExplicitLetStatement,
                isSetAssignment);
        }
    }
}
