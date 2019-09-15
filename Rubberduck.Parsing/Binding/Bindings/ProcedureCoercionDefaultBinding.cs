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
        private readonly bool _hasExplicitCall;
        private readonly Declaration _parent;
        private IBoundExpression _wrappedExpression;

        //This is a wrapper used to model procedure coercion in call statements without arguments.
        //The one with arguments is basically an index expression and uses its binding.

        public ProcedureCoercionDefaultBinding(
            ParserRuleContext expression,
            IExpressionBinding wrappedExpressionBinding,
            bool hasExplicitCall,
            Declaration parent)
            : this(
                expression,
                (IBoundExpression)null,
                hasExplicitCall,
                parent)
        {
            _wrappedExpressionBinding = wrappedExpressionBinding;
        }

        public ProcedureCoercionDefaultBinding(
            ParserRuleContext expression,
            IBoundExpression wrappedExpression,
            bool hasExplicitCall,
            Declaration parent)
        {
            _expression = expression;
            _wrappedExpression = wrappedExpression;
            _hasExplicitCall = hasExplicitCall;
            _parent = parent;
        }

        public IBoundExpression Resolve()
        {
            if (_wrappedExpressionBinding != null)
            {
                _wrappedExpression = _wrappedExpressionBinding.Resolve();
            }

            return Resolve(_wrappedExpression, _expression, _hasExplicitCall, _parent);
        }

        private static IBoundExpression Resolve(IBoundExpression wrappedExpression, ParserRuleContext expression, bool hasExplicitCall, Declaration parent)
        {
            //Procedure coercion only happens for expressions classified as variables.
            if (wrappedExpression.Classification != ExpressionClassification.Variable)
            {
                return wrappedExpression;
            }

            //The wrapped declaration is not of a specific class type or Object.
            var wrappedDeclaration = wrappedExpression.ReferencedDeclaration;
            if (wrappedDeclaration == null
                || !wrappedDeclaration.IsObject
                    && !(wrappedDeclaration.IsObjectArray
                        && wrappedExpression is IndexExpression arrayExpression
                        && arrayExpression.IsArrayAccess))
            {
                return wrappedExpression;
            }

            //Recursive function call
            //The reference to the function is originally resolved as a variable because that is appropriate for the return value variable of the same name.
            if (wrappedExpression.Classification == ExpressionClassification.Variable
                && wrappedDeclaration.Equals(parent))
            {
                return wrappedExpression;
            }

            var asTypeName = wrappedDeclaration.AsTypeName;
            var asTypeDeclaration = wrappedDeclaration.AsTypeDeclaration;

            //If there is an explicit call, a non-array (access) index expression or dictionary access expression already count as procedure call.
            if (hasExplicitCall
                && (wrappedExpression is IndexExpression indexExpression 
                        && !indexExpression.IsArrayAccess
                    || wrappedExpression is DictionaryAccessExpression))
            {
                return wrappedExpression;
            }

            return ResolveViaDefaultMember(wrappedExpression, asTypeName, asTypeDeclaration, expression);
        }

        private static IBoundExpression CreateFailedExpression(IBoundExpression wrappedExpression, ParserRuleContext context)
        {
            return new ProcedureCoercionExpression(wrappedExpression.ReferencedDeclaration, ExpressionClassification.ResolutionFailed, context, wrappedExpression);
        }

        private static IBoundExpression ResolveViaDefaultMember(IBoundExpression wrappedExpression, string asTypeName, Declaration asTypeDeclaration, ParserRuleContext expression)
        {
            if (Tokens.Variant.Equals(asTypeName, StringComparison.InvariantCultureIgnoreCase)
                    || Tokens.Object.Equals(asTypeName, StringComparison.InvariantCultureIgnoreCase))
            {
                // We cannot know the the default member in this case, so return an unbound member call.
                return new ProcedureCoercionExpression(wrappedExpression.ReferencedDeclaration, ExpressionClassification.Unbound, expression, wrappedExpression);
            }

            var defaultMember = (asTypeDeclaration as ClassModuleDeclaration)?.DefaultMember;
            if (defaultMember == null
                || !IsPropertyGetLetFunctionProcedure(defaultMember)
                || !IsPublic(defaultMember))
            {
                return CreateFailedExpression(wrappedExpression, expression);
            }

            var defaultMemberClassification = DefaultMemberClassification(defaultMember);

            var parameters = ((IParameterizedDeclaration)defaultMember).Parameters.ToList();
            if (parameters.All(parameter => parameter.IsOptional || parameter.IsParamArray))
            {
                //We found some default member accepting the empty argument list. So, we are done.
                return new ProcedureCoercionExpression(defaultMember, defaultMemberClassification, expression, wrappedExpression);
            }

            return CreateFailedExpression(wrappedExpression, expression);
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