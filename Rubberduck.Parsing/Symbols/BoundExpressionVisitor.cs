using Antlr4.Runtime;
using Rubberduck.Parsing.Binding;
using System;

namespace Rubberduck.Parsing.Symbols
{
    public sealed class BoundExpressionVisitor
    {
        public void AddIdentifierReferences(IBoundExpression boundExpression, Func<ParserRuleContext, Declaration, IdentifierReference> referenceCreator)
        {
            Visit((dynamic)boundExpression, referenceCreator);
        }

        private void Visit(SimpleNameExpression expression, Func<ParserRuleContext, Declaration, IdentifierReference> referenceCreator)
        {
            expression.ReferencedDeclaration.AddReference(referenceCreator(expression.Context, expression.ReferencedDeclaration));
        }

        private void Visit(MemberAccessExpression expression, Func<ParserRuleContext, Declaration, IdentifierReference> referenceCreator)
        {
            Visit((dynamic)expression.LExpression, referenceCreator);
            // Expressions could be unbound thus not have a referenced declaration. The lexpression might still be bindable though.
            if (expression.Classification != ExpressionClassification.Unbound)
            {
                expression.ReferencedDeclaration.AddReference(referenceCreator(expression.Context, expression.ReferencedDeclaration));
            }
        }

        private void Visit(IndexExpression expression, Func<ParserRuleContext, Declaration, IdentifierReference> referenceCreator)
        {
            Visit((dynamic)expression.LExpression, referenceCreator);
            // Expressions could be unbound thus not have a referenced declaration. The lexpression might still be bindable though.
            if (expression.Classification != ExpressionClassification.Unbound)
            {
                // Referenced declaration could also be null if e.g. it's an array and the array is a "base type" such as String.
                if (expression.ReferencedDeclaration != null)
                {
                    expression.ReferencedDeclaration.AddReference(referenceCreator(expression.Context, expression.ReferencedDeclaration));
                }
            }
            // Argument List not affected by being unbound.
            foreach (var argument in expression.ArgumentList.Arguments)
            {
                if (argument.Expression != null)
                {
                    Visit((dynamic)argument.Expression, referenceCreator);
                }
                if (argument.NamedArgumentExpression != null)
                {
                    Visit((dynamic)argument.NamedArgumentExpression, referenceCreator);
                }
            }
        }

        private void Visit(NewExpression expression, Func<ParserRuleContext, Declaration, IdentifierReference> referenceCreator)
        {
            // We don't need to add a reference to the NewExpression's referenced declaration since that's covered
            // with its TypeExpression.
            Visit((dynamic)expression.TypeExpression, referenceCreator);
        }

        private void Visit(ParenthesizedExpression expression, Func<ParserRuleContext, Declaration, IdentifierReference> referenceCreator)
        {
            Visit((dynamic)expression.Expression, referenceCreator);
        }

        private void Visit(TypeOfIsExpression expression, Func<ParserRuleContext, Declaration, IdentifierReference> referenceCreator)
        {
            Visit((dynamic)expression.Expression, referenceCreator);
            Visit((dynamic)expression.TypeExpression, referenceCreator);
        }

        private void Visit(BinaryOpExpression expression, Func<ParserRuleContext, Declaration, IdentifierReference> referenceCreator)
        {
            Visit((dynamic)expression.Left, referenceCreator);
            Visit((dynamic)expression.Right, referenceCreator);
        }

        private void Visit(UnaryOpExpression expression, Func<ParserRuleContext, Declaration, IdentifierReference> referenceCreator)
        {
            Visit((dynamic)expression.Expr, referenceCreator);
        }

        private void Visit(LiteralExpression expression, Func<ParserRuleContext, Declaration, IdentifierReference> referenceCreator)
        {
            // Nothing to do here.
        }

        private void Visit(InstanceExpression expression, Func<ParserRuleContext, Declaration, IdentifierReference> referenceCreator)
        {
            expression.ReferencedDeclaration.AddReference(referenceCreator(expression.Context, expression.ReferencedDeclaration));
        }
    }
}
