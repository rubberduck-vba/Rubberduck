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
            expression.ReferencedDeclaration.AddReference(referenceCreator(expression.ReferencedDeclaration));
        }
    }
}
