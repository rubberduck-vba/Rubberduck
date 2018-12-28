using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Binding;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.VBEditor;

// ReSharper disable UnusedParameter.Local  - calls are dynamic, so the signatures need to match.

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
            Visit((dynamic)boundExpression, module, scope, parent, isAssignmentTarget, hasExplicitLetStatement, isSetAssignment);
        }

        private void Visit(
            ResolutionFailedExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent,
            bool isAssignmentTarget,
            bool hasExplicitLetStatement,
            bool isSetAssignment)
        {
            // To bind as much as possible we gather all successfully resolved expressions and bind them here as a special case.
            foreach (var successfullyResolvedExpression in expression.SuccessfullyResolvedExpressions)
            {
                Visit((dynamic)successfullyResolvedExpression, module, scope, parent, false, false, false);
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
            Visit((dynamic)expression.LExpression, module, scope, parent, false, false, false);
            // Expressions could be unbound thus not have a referenced declaration. The lexpression might still be bindable though.
            if (expression.Classification != ExpressionClassification.Unbound)
            {
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
            Visit((dynamic)expression.LExpression, module, scope, parent, isAssignmentTarget, hasExplicitLetStatement, isSetAssignment);
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
                    Visit((dynamic)argument.Expression, module, scope, parent, false, false, false);
                }
                if (argument.NamedArgumentExpression != null)
                {
                    Visit((dynamic)argument.NamedArgumentExpression, module, scope, parent, false, false, false);
                }
            }
        }

        private void Visit(
            NewExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent,
            bool isAssignmentTarget,
            bool hasExplicitLetStatement,
            bool isSetAssignment)
        {
            // We don't need to add a reference to the NewExpression's referenced declaration since that's covered
            // with its TypeExpression.
            Visit((dynamic)expression.TypeExpression, module, scope, parent, false, false, false);
        }

        private void Visit(
            ParenthesizedExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent,
            bool isAssignmentTarget,
            bool hasExplicitLetStatement,
            bool isSetAssignment)
        {
            Visit((dynamic)expression.Expression, module, scope, parent, false, false, false);
        }

        private void Visit(
            TypeOfIsExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent,
            bool isAssignmentTarget,
            bool hasExplicitLetStatement,
            bool isSetAssignment)
        {
            Visit((dynamic)expression.Expression, module, scope, parent, false, false, false);
            Visit((dynamic)expression.TypeExpression, module, scope, parent, false, false, false);
        }

        private void Visit(
            BinaryOpExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent,
            bool isAssignmentTarget,
            bool hasExplicitLetStatement,
            bool isSetAssignment)
        {
            Visit((dynamic)expression.Left, module, scope, parent, false, false, false);
            Visit((dynamic)expression.Right, module, scope, parent, false, false, false);
        }

        private void Visit(
            UnaryOpExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent,
            bool isAssignmentTarget,
            bool hasExplicitLetStatement,
            bool isSetAssignment)
        {
            Visit((dynamic)expression.Expr, module, scope, parent, false, false, false);
        }

        private void Visit(
            LiteralExpression expression,
            QualifiedModuleName module,
            Declaration scope,
            Declaration parent,
            bool isAssignmentTarget,
            bool hasExplicitLetStatement,
            bool isSetAssignment)
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
