using Rubberduck.Parsing.Binding;
using System;

namespace Rubberduck.Parsing.Symbols
{
    public sealed class BoundExpressionVisitor
    {
        public void AddIdentifierReferences(IBoundExpression boundExpression, Func<Declaration, IdentifierReference> referenceCreator)
        {
            Visit((dynamic)boundExpression, referenceCreator);
        }

        private void Visit(SimpleNameExpression expression, Func<Declaration, IdentifierReference> referenceCreator)
        {
            expression.ReferencedDeclaration.AddReference(referenceCreator(expression.ReferencedDeclaration));
        }

        private void Visit(MemberAccessExpression expression, Func<Declaration, IdentifierReference> referenceCreator)
        {
            Visit((dynamic)expression.LExpression, referenceCreator);
            // Expressions could be unbound thus not have a referenced declaration. The lexpression might still be bindable though.
            if (expression.Classification != ExpressionClassification.Unbound)
            {
                expression.ReferencedDeclaration.AddReference(referenceCreator(expression.ReferencedDeclaration));
            }
        }

        private void Visit(IndexExpression expression, Func<Declaration, IdentifierReference> referenceCreator)
        {
            Visit((dynamic)expression.LExpression, referenceCreator);
            // Expressions could be unbound thus not have a referenced declaration. The lexpression might still be bindable though.
            if (expression.Classification != ExpressionClassification.Unbound)
            {
                expression.ReferencedDeclaration.AddReference(referenceCreator(expression.ReferencedDeclaration));
            }
            // Argument List not affected by being unbound.
            foreach (var argument in expression.ArgumentList.Arguments)
            {
                Visit((dynamic)argument.Expression, referenceCreator);
            }
        }

        private void Visit(NewExpression expression, Func<Declaration, IdentifierReference> referenceCreator)
        {
            // We don't need to add a reference to the NewExpression's referenced declaration since that's covered
            // with its TypeExpression.
            Visit((dynamic)expression.TypeExpression, referenceCreator);
        }

        private void Visit(ParenthesizedExpression expression, Func<Declaration, IdentifierReference> referenceCreator)
        {
            Visit((dynamic)expression.Expression, referenceCreator);
        }

        private void Visit(TypeOfIsExpression expression, Func<Declaration, IdentifierReference> referenceCreator)
        {
            Visit((dynamic)expression.Expression, referenceCreator);
            Visit((dynamic)expression.TypeExpression, referenceCreator);
        }

        private void Visit(BinaryOpExpression expression, Func<Declaration, IdentifierReference> referenceCreator)
        {
            Visit((dynamic)expression.Left, referenceCreator);
            Visit((dynamic)expression.Right, referenceCreator);
        }

        private void Visit(UnaryOpExpression expression, Func<Declaration, IdentifierReference> referenceCreator)
        {
            Visit((dynamic)expression.Expr, referenceCreator);
        }

        private void Visit(LiteralExpression expression, Func<Declaration, IdentifierReference> referenceCreator)
        {
            // Nothing to do here.
        }

        private void Visit(InstanceExpression expression, Func<Declaration, IdentifierReference> referenceCreator)
        {
            expression.ReferencedDeclaration.AddReference(referenceCreator(expression.ReferencedDeclaration));
        }
    }
}
