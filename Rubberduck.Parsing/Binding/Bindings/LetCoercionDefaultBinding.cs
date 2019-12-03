using System;
using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.Binding
{
    public sealed class LetCoercionDefaultBinding : IExpressionBinding
    {
        private readonly ParserRuleContext _expression;
        private readonly IExpressionBinding _wrappedExpressionBinding;
        private IBoundExpression _wrappedExpression;
        private readonly bool _isAssignment;

        private const int DEFAULT_MEMBER_RECURSION_LIMIT = 32;

        //This is a wrapper used to model Let coercion for object types.

        public LetCoercionDefaultBinding(
            ParserRuleContext expression,
            IExpressionBinding wrappedExpressionBinding,
            bool isAssignment = false)
            : this(
                expression,
                (IBoundExpression)null,
                isAssignment)
        {
            _wrappedExpressionBinding = wrappedExpressionBinding;
        }

        public LetCoercionDefaultBinding(
            ParserRuleContext expression,
            IBoundExpression wrappedExpression,
            bool isAssignment = false)
        {
            _expression = expression;
            _wrappedExpression = wrappedExpression;
            _isAssignment = isAssignment;
        }

        public IBoundExpression Resolve()
        {
            if (_wrappedExpressionBinding != null)
            {
                _wrappedExpression = _wrappedExpressionBinding.Resolve();
            }

            return Resolve(_wrappedExpression, _expression, _isAssignment);
        }

        private static IBoundExpression Resolve(IBoundExpression wrappedExpression, ParserRuleContext expression, bool isAssignment)
        {
            if (wrappedExpression.Classification == ExpressionClassification.ResolutionFailed)
            {
                return wrappedExpression;
            }

            var wrappedDeclaration = wrappedExpression.ReferencedDeclaration;

            if (wrappedDeclaration == null
                || !wrappedDeclaration.IsObject 
                    && !(wrappedDeclaration.IsObjectArray 
                        && wrappedExpression is IndexExpression indexExpression 
                        && indexExpression.IsArrayAccess)
                || wrappedDeclaration.DeclarationType == DeclarationType.PropertyLet)
            {
                return wrappedExpression;
            }

            //The wrapped declaration is of a specific class type or Object.

            if (wrappedExpression.Classification == ExpressionClassification.Unbound)
            {
                //This should actually not be possible since an unbound expression cannot have a referenced declaration. 
                //Apart from this, we can only deal with the type Object.
                return new LetCoercionDefaultMemberAccessExpression(null, ExpressionClassification.Unbound, expression, wrappedExpression, 1, null);
            }

            var asTypeName = wrappedDeclaration.AsTypeName;
            var asTypeDeclaration = wrappedDeclaration.AsTypeDeclaration;

            return ResolveViaDefaultMember(wrappedExpression, asTypeName, asTypeDeclaration, expression, isAssignment);
        }

        private static IBoundExpression ExpressionForResolutionFailure(IBoundExpression wrappedExpression, ParserRuleContext expression)
        {
            //We return a LetCoercionExpression classified as failed to enable us to save this failed coercion.
            return new LetCoercionDefaultMemberAccessExpression(wrappedExpression.ReferencedDeclaration, ExpressionClassification.ResolutionFailed, expression, wrappedExpression, 1, null);
        }

        private static IBoundExpression ResolveViaDefaultMember(IBoundExpression wrappedExpression, string asTypeName, Declaration asTypeDeclaration, ParserRuleContext expression, bool isAssignment, int recursionDepth = 1, RecursiveDefaultMemberAccessExpression containedExpression = null)
        {
            if (Tokens.Variant.Equals(asTypeName, StringComparison.InvariantCultureIgnoreCase)
                    || Tokens.Object.Equals(asTypeName, StringComparison.InvariantCultureIgnoreCase))
            {
                // We cannot know the the default member in this case, so return an unbound member call.
                return new LetCoercionDefaultMemberAccessExpression(wrappedExpression.ReferencedDeclaration, ExpressionClassification.Unbound, expression, wrappedExpression, recursionDepth, containedExpression);
            }

            var defaultMember = (asTypeDeclaration as ClassModuleDeclaration)?.DefaultMember;
            if (defaultMember == null
                || !IsPropertyGetLetFunctionProcedure(defaultMember)
                || !IsPublic(defaultMember))
            {
                return ExpressionForResolutionFailure(wrappedExpression, expression);
            }

            var defaultMemberClassification = DefaultMemberClassification(defaultMember);

            var parameters = ((IParameterizedDeclaration)defaultMember).Parameters.ToList();
            if (isAssignment
                && defaultMember.DeclarationType == DeclarationType.PropertyLet
                && IsCompatibleWithOneNonObjectParameter(parameters))
            {
                //This is a Let assignment. So, finding a Property Let with one non object parameter means we are done.
                return new LetCoercionDefaultMemberAccessExpression(defaultMember, defaultMemberClassification, expression, wrappedExpression, recursionDepth, containedExpression);
            }

            if (parameters.All(parameter => parameter.IsOptional || parameter.IsParamArray))
            {
                if (!defaultMember.IsObject)
                {
                    //We found a property Get of Function default member returning a value type.
                    //This might also be applicable in case of an assignment, because only the Get will be assigned as default member if both Get and Let exist.
                    return new LetCoercionDefaultMemberAccessExpression(defaultMember, defaultMemberClassification, expression, wrappedExpression, recursionDepth, containedExpression);
                }

                if (DEFAULT_MEMBER_RECURSION_LIMIT >= recursionDepth)
                {
                    //The default member returns an object type. So, we have to recurse.
                    return ResolveRecursiveDefaultMember(wrappedExpression, defaultMember, defaultMemberClassification, expression, isAssignment, recursionDepth, containedExpression);
                }
            }

            return ExpressionForResolutionFailure(wrappedExpression, expression);
        }

        private static IBoundExpression ResolveRecursiveDefaultMember(IBoundExpression wrappedExpression, Declaration defaultMember, ExpressionClassification defaultMemberClassification, ParserRuleContext expression, bool isAssignment, int recursionDepth, RecursiveDefaultMemberAccessExpression containedExpression)
        {
            var defaultMemberAsTypeName = defaultMember.AsTypeName;
            var defaultMemberAsTypeDeclaration = defaultMember.AsTypeDeclaration;

            var defaultMemberExpression = new RecursiveDefaultMemberAccessExpression(defaultMember, defaultMemberClassification, expression, recursionDepth, containedExpression);

            return ResolveViaDefaultMember(
                wrappedExpression,
                defaultMemberAsTypeName,
                defaultMemberAsTypeDeclaration,
                expression,
                isAssignment,
                recursionDepth + 1,
                defaultMemberExpression);
        }

        private static bool IsCompatibleWithOneNonObjectParameter(IReadOnlyCollection<ParameterDeclaration> parameters)
        {
            return parameters.Count(parameter => !parameter.IsObject) == 1
                        && parameters.Any(parameter => !parameter.IsOptional && !parameter.IsObject)
                    || parameters.All(parameter => parameter.IsOptional)
                        && parameters.Any(parameter => !parameter.IsObject);
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