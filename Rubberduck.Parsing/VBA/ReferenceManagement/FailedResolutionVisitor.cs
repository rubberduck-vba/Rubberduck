using System;
using NLog;
using Rubberduck.Parsing.Binding;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA.DeclarationCaching;

namespace Rubberduck.Parsing.VBA.ReferenceManagement
{
    public sealed class FailedResolutionVisitor
    {
        private readonly DeclarationFinder _declarationFinder;

        private static Logger Logger = LogManager.GetCurrentClassLogger();

        public FailedResolutionVisitor(DeclarationFinder declarationFinder)
        {
            _declarationFinder = declarationFinder;
        }

        public void CollectUnresolved(IBoundExpression boundExpression, Declaration parent, IBoundExpression withExpression)
        {
            Visit(boundExpression, parent, withExpression);
        }

        private void Visit(IBoundExpression boundExpression, Declaration parent, IBoundExpression withExpression)
        {
            switch (boundExpression)
            {
                case SimpleNameExpression simpleNameExpression:
                    break;
                case MemberAccessExpression memberAccessExpression:
                    Visit(memberAccessExpression, parent, withExpression);
                    break;
                case IndexExpression indexExpression:
                    Visit(indexExpression, parent, withExpression);
                    break;
                case ParenthesizedExpression parenthesizedExpression:
                    Visit(parenthesizedExpression, parent, withExpression);
                    break;
                case LiteralExpression literalExpression:
                    break;
                case BinaryOpExpression binaryOpExpression:
                    Visit(binaryOpExpression, parent, withExpression);
                    break;
                case UnaryOpExpression unaryOpExpression:
                    Visit(unaryOpExpression, parent, withExpression);
                    break;
                case NewExpression newExpression:
                    Visit(newExpression, parent, withExpression);
                    break;
                case InstanceExpression instanceExpression:
                    break;
                case DictionaryAccessExpression dictionaryAccessExpression:
                    Visit(dictionaryAccessExpression, parent, withExpression);
                    break;
                case TypeOfIsExpression typeOfIsExpression:
                    Visit(typeOfIsExpression, parent, withExpression);
                    break;
                case ResolutionFailedExpression resolutionFailedExpression:
                    Visit(resolutionFailedExpression, parent, withExpression);
                    break;
                case BuiltInTypeExpression builtInTypeExpression:
                    break;
                case RecursiveDefaultMemberAccessExpression recursiveDefaultMemberAccessExpression:
                    break;
                case LetCoercionDefaultMemberAccessExpression letCoercionDefaultMemberAccessExpression:
                    Visit(letCoercionDefaultMemberAccessExpression, parent, withExpression);
                    break;
                case ProcedureCoercionExpression procedureCoercionExpression:
                    Visit(procedureCoercionExpression, parent, withExpression);
                    break;
                case OutputListExpression outputListExpression:
                    Visit(outputListExpression, parent, withExpression);
                    break;
                case ObjectPrintExpression objectPrintExpression:
                    Visit(objectPrintExpression, parent, withExpression);
                    break;
                case MissingArgumentExpression missingArgumentExpression:
                    break;
                default:
                    throw new NotSupportedException($"Unexpected bound expression type {boundExpression.GetType()}");
            }
        }

        private void Visit(ResolutionFailedExpression expression, Declaration parent, IBoundExpression withExpression)
        {
            if (!expression.IsJoinedExpression)
            {
                SaveUnresolvedExpression(expression, parent, withExpression);
            }

            foreach (var successfullyResolvedExpression in expression.SuccessfullyResolvedExpressions)
            {
                Visit(successfullyResolvedExpression, parent, withExpression);
            }
        }

        private void SaveUnresolvedExpression(ResolutionFailedExpression expression, Declaration parent, IBoundExpression withExpression)
        {
            if (expression.Context is VBAParser.LExpressionContext lExpression)
            {
                _declarationFinder.AddUnboundContext(parent, lExpression, withExpression);
            }
            else
            {
                Logger.Warn($"Default Context: Failed to resolve {expression.Context.GetText()}. Binding as much as we can.");
            }
        }

        private void Visit(MemberAccessExpression expression, Declaration parent, IBoundExpression withExpression)
        {
            Visit(expression.LExpression, parent, withExpression);
        }

        private void Visit(ObjectPrintExpression expression, Declaration parent, IBoundExpression withExpression)
        {
            var outputListExpression = expression.OutputListExpression;
            if (outputListExpression != null)
            {
                Visit(expression.OutputListExpression, parent, withExpression);
            }

            Visit(expression.PrintMethodExpressions, parent, withExpression);
        }

        private void Visit(OutputListExpression expression, Declaration parent, IBoundExpression withExpression)
        {
            foreach (var itemExpression in expression.ItemExpressions)
            {
                Visit(itemExpression, parent, withExpression);
            }
        }

        private void Visit(IndexExpression expression, Declaration parent, IBoundExpression withExpression)
        {
            Visit(expression.LExpression, parent, withExpression);

            foreach (var argument in expression.ArgumentList.Arguments)
            {
                if (argument.Expression != null)
                {
                    Visit(argument.Expression, parent, withExpression);
                }
            }
        }

        private void Visit(DictionaryAccessExpression expression, Declaration parent, IBoundExpression withExpression)
        {
            Visit(expression.LExpression, parent, withExpression);
        }

        private void Visit(LetCoercionDefaultMemberAccessExpression expression, Declaration parent, IBoundExpression withExpression)
        {
            Visit(expression.WrappedExpression, parent, withExpression);
        }

        private void Visit(ProcedureCoercionExpression expression, Declaration parent, IBoundExpression withExpression)
        {
            Visit(expression.WrappedExpression, parent, withExpression);
        }

        private void Visit(NewExpression expression, Declaration parent, IBoundExpression withExpression)
        {
            Visit(expression.TypeExpression, parent, withExpression);
        }

        private void Visit(ParenthesizedExpression expression, Declaration parent, IBoundExpression withExpression)
        {
            Visit(expression.Expression, parent, withExpression);
        }

        private void Visit(TypeOfIsExpression expression, Declaration parent, IBoundExpression withExpression)
        {
            Visit(expression.Expression, parent, withExpression);
            Visit(expression.TypeExpression, parent, withExpression);
        }

        private void Visit(BinaryOpExpression expression, Declaration parent, IBoundExpression withExpression)
        {
            Visit(expression.Left, parent, withExpression);
            Visit(expression.Right, parent, withExpression);
        }

        private void Visit(UnaryOpExpression expression, Declaration parent, IBoundExpression withExpression)
        {
            Visit(expression.Expr, parent, withExpression);
        }
    }
}