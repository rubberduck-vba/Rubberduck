using System;
using Rubberduck.Parsing.Symbols;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.Binding
{
    public sealed class ProcedureCoercionDefaultBinding : IExpressionBinding
    {
        private readonly ParserRuleContext _expression;
        private readonly IExpressionBinding _wrappedExpressionBinding;
        private IBoundExpression _wrappedExpression;

        //This is a wrapper used to model procedure coercion in call statements without arguments.
        //The one with arguments is basically an index expression and uses its binding.

        public ProcedureCoercionDefaultBinding(
            ParserRuleContext expression,
            IExpressionBinding wrappedExpressionBinding)
            : this(
                expression,
                (IBoundExpression)null)
        {
            _wrappedExpressionBinding = wrappedExpressionBinding;
        }

        public ProcedureCoercionDefaultBinding(
            ParserRuleContext expression,
            IBoundExpression wrappedExpression)
        {
            _expression = expression;
            _wrappedExpression = wrappedExpression;
        }

        public IBoundExpression Resolve()
        {
            if (_wrappedExpressionBinding != null)
            {
                _wrappedExpression = _wrappedExpressionBinding.Resolve();
            }

            return Resolve(_wrappedExpression, _expression);
        }

        private static IBoundExpression Resolve(IBoundExpression wrappedExpression, ParserRuleContext expression)
        {
            //Procedure coercion only happens for expressions classified as variables.
            if (wrappedExpression.Classification != ExpressionClassification.Variable)
            {
                return wrappedExpression;
            }

            var wrappedDeclaration = wrappedExpression.ReferencedDeclaration;
            if (wrappedDeclaration == null
                || !wrappedDeclaration.IsObject
                    && !(wrappedDeclaration.IsObjectArray
                        && wrappedExpression is IndexExpression indexExpression
                        && indexExpression.IsArrayAccess))
            {
                return wrappedExpression;
            }

            //The wrapped declaration is of a specific class type or Object.

            var asTypeName = wrappedDeclaration.AsTypeName;
            var asTypeDeclaration = wrappedDeclaration.AsTypeDeclaration;

            return ResolveViaDefaultMember(wrappedExpression, asTypeName, asTypeDeclaration, expression);
        }

        private static IBoundExpression CreateFailedExpression(IBoundExpression lExpression)
        {
            var failedExpr = new ResolutionFailedExpression();
            failedExpr.AddSuccessfullyResolvedExpression(lExpression);
            return failedExpr;
        }

        private static IBoundExpression ResolveViaDefaultMember(IBoundExpression wrappedExpression, string asTypeName, Declaration asTypeDeclaration, ParserRuleContext expression)
        {
            if (Tokens.Variant.Equals(asTypeName, StringComparison.InvariantCultureIgnoreCase)
                    || Tokens.Object.Equals(asTypeName, StringComparison.InvariantCultureIgnoreCase))
            {
                // We cannot know the the default member in this case, so return an unbound member call.
                return new ProcedureCoercionExpression(null, ExpressionClassification.Unbound, expression, wrappedExpression);
            }

            var defaultMember = (asTypeDeclaration as ClassModuleDeclaration)?.DefaultMember;
            if (defaultMember == null
                || !IsPropertyGetLetFunctionProcedure(defaultMember)
                || !IsPublic(defaultMember))
            {
                return CreateFailedExpression(wrappedExpression);
            }

            var defaultMemberClassification = DefaultMemberClassification(defaultMember);

            var parameters = ((IParameterizedDeclaration)defaultMember).Parameters.ToList();
            if (parameters.All(parameter => parameter.IsOptional))
            {
                //We found some default member accepting the empty argument list. So, we are done.
                return new ProcedureCoercionExpression(defaultMember, defaultMemberClassification, expression, wrappedExpression);
            }

            return CreateFailedExpression(wrappedExpression);
        }

        private static bool IsPropertyGetLetFunctionProcedure(Declaration declaration)
        {
            var declarationType = declaration.DeclarationType;
            return declarationType == DeclarationType.PropertyGet
                || declarationType == DeclarationType.PropertyLet
                || declarationType == DeclarationType.Function
                || declarationType == DeclarationType.Procedure;
        }

        private static bool IsPublic(Declaration declaration)
        {
            var accessibility = declaration.Accessibility;
            return accessibility == Accessibility.Global
                   || accessibility == Accessibility.Implicit
                   || accessibility == Accessibility.Public;
        }

        private static ExpressionClassification DefaultMemberClassification(Declaration defaultMember)
        {
            if (defaultMember.DeclarationType.HasFlag(DeclarationType.Property))
            {
                return ExpressionClassification.Property;
            }

            if (defaultMember.DeclarationType == DeclarationType.Procedure)
            {
                return ExpressionClassification.Subroutine;
            }

            return ExpressionClassification.Function;
        }
    }
}