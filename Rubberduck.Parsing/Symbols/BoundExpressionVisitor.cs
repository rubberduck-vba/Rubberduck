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

        private void Visit(NewExpression expression, Func<Declaration, IdentifierReference> referenceCreator)
        {
            // We don't need to add a reference to the NewExpression's referenced declaration since that's covered
            // with its TypeExpression.
            Visit((dynamic)expression.TypeExpression, referenceCreator);
        }
    }
}
