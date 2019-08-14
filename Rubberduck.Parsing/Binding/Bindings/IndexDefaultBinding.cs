using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;
using System.Linq;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.Binding
{
    public sealed class IndexDefaultBinding : IExpressionBinding
    {
        private readonly ParserRuleContext _expression;
        private readonly IExpressionBinding _lExpressionBinding;
        private IBoundExpression _lExpression;
        private readonly ArgumentList _argumentList;

        private const int DEFAULT_MEMBER_RECURSION_LIMIT = 32;

        public IndexDefaultBinding(
            ParserRuleContext expression,
            IExpressionBinding lExpressionBinding,
            ArgumentList argumentList)
            : this(
                  expression,
                  (IBoundExpression)null,
                  argumentList)
        {
            _lExpressionBinding = lExpressionBinding;
        }

        public IndexDefaultBinding(
            ParserRuleContext expression,
            IBoundExpression lExpression,
            ArgumentList argumentList)
        {
            _expression = expression;
            _lExpression = lExpression;
            _argumentList = argumentList;
        }

        private void ResolveArgumentList(Declaration calledProcedure)
        {
            foreach (var argument in _argumentList.Arguments)
            {
                argument.Resolve(calledProcedure);
            }
        }

        public IBoundExpression Resolve()
        {
            if (_lExpressionBinding != null)
            {
                _lExpression = _lExpressionBinding.Resolve();
            }
            if (_lExpression.Classification != ExpressionClassification.ResolutionFailed)
            {
                ResolveArgumentList(_lExpression.ReferencedDeclaration);
            }
            else
            {
                ResolveArgumentList(null);
            }
            return Resolve(_lExpression);
        }

        private IBoundExpression Resolve(IBoundExpression lExpression, int defaultMemberResolutionRecursionDepth = 0)
        {
            if (lExpression.Classification == ExpressionClassification.ResolutionFailed)
            {
                return CreateFailedExpression(lExpression);
            }

            if (lExpression.Classification == ExpressionClassification.Unbound)
            {
                return ResolveLExpressionIsUnbound(lExpression);
            }

            if(lExpression.ReferencedDeclaration != null)
            {
                if (_argumentList.HasArguments)
                {
                    switch (lExpression)
                    {
                        case IndexExpression indexExpression:
                            var doubleIndexExpression = ResolveLExpressionIsIndexExpression(indexExpression, defaultMemberResolutionRecursionDepth);
                            if (doubleIndexExpression != null)
                            {
                                return doubleIndexExpression;
                            }

                            break;
                        case DictionaryAccessExpression dictionaryAccessExpression:
                            var indexOnBangExpression = ResolveLExpressionIsDictionaryAccessExpression(dictionaryAccessExpression, defaultMemberResolutionRecursionDepth);
                            if (indexOnBangExpression != null)
                            {
                                return indexOnBangExpression;
                            }

                            break;
                    }
                }

                if (IsVariablePropertyFunctionWithoutParameters(lExpression))
                {
                    var parameterlessLExpressionAccess = ResolveLExpressionIsVariablePropertyFunctionNoParameters(lExpression, defaultMemberResolutionRecursionDepth);
                    if (parameterlessLExpressionAccess != null)
                    {
                        return parameterlessLExpressionAccess;
                    }
                }
            }

            if (lExpression.Classification == ExpressionClassification.Property
                || lExpression.Classification == ExpressionClassification.Function
                || lExpression.Classification == ExpressionClassification.Subroutine)
            {
                return ResolveLExpressionIsPropertyFunctionSubroutine(lExpression);
            }

            return CreateFailedExpression(lExpression);
        }

        private IBoundExpression CreateFailedExpression(IBoundExpression lExpression)
        {
            var failedExpr = new ResolutionFailedExpression();
            failedExpr.AddSuccessfullyResolvedExpression(lExpression);
            foreach (var arg in _argumentList.Arguments)
            {
                failedExpr.AddSuccessfullyResolvedExpression(arg.Expression);
            }
            return failedExpr;
        }

        private IBoundExpression ResolveLExpressionIsVariablePropertyFunctionNoParameters(IBoundExpression lExpression, int defaultMemberResolutionRecursionDepth)
        {
            /*
                <l-expression> is classified as a variable, or <l-expression> is classified as a property or function 
                with a parameter list that cannot accept any parameters and an <argument-list> that is not 
                empty, and one of the following is true (see below):

                There are no parameters to the lExpression. So, this is either an array access or a default member call.
             */

            var indexedDeclaration = lExpression.ReferencedDeclaration;
            if (indexedDeclaration == null)
            {
                return null;
            }

            if (indexedDeclaration.IsArray)
            {
                return ResolveLExpressionDeclaredTypeIsArray(lExpression);
            }

            var asTypeName = indexedDeclaration.AsTypeName;
            var asTypeDeclaration = indexedDeclaration.AsTypeDeclaration;

            return ResolveDefaultMember(lExpression, asTypeName, asTypeDeclaration, defaultMemberResolutionRecursionDepth);
        }

        private static bool IsVariablePropertyFunctionWithoutParameters(IBoundExpression lExpression)
        {
            switch(lExpression.Classification)
            {
                case ExpressionClassification.Variable:
                    return true;
                case ExpressionClassification.Function:
                case ExpressionClassification.Property:
                    return !((IParameterizedDeclaration)lExpression.ReferencedDeclaration).Parameters.Any();
                default:
                    return false;
            }
        }

        private IBoundExpression ResolveLExpressionIsIndexExpression(IndexExpression indexExpression, int defaultMemberResolutionRecursionDepth = 0)
        {
            /*
             <l-expression> is classified as an index expression and the argument list is not empty.
                Thus, me must be dealing with a default member access or an array access.
             */

            var indexedDeclaration = indexExpression.ReferencedDeclaration;
            if (indexedDeclaration == null)
            {
                return null;
            }

            //The result of an array access is never an array. Any double array access requires either a default member access in between
            //or an array assigned to a Variant, the access to which is counted as an unbound member access and, thus, is resolved correctly
            //via the default member path.
            if (indexedDeclaration.IsArray && !indexExpression.IsArrayAccess)
            {
                return ResolveLExpressionDeclaredTypeIsArray(indexExpression);
            }

            var asTypeName = indexedDeclaration.AsTypeName;
            var asTypeDeclaration = indexedDeclaration.AsTypeDeclaration;

            return ResolveDefaultMember(indexExpression, asTypeName, asTypeDeclaration, defaultMemberResolutionRecursionDepth);
        }

        private IBoundExpression ResolveLExpressionIsDictionaryAccessExpression(DictionaryAccessExpression dictionaryAccessExpression, int defaultMemberResolutionRecursionDepth = 0)
        {
            //This is equivalent to the case in which the lExpression is an IndexExpression with the difference that it cannot be an array access.

            var indexedDeclaration = dictionaryAccessExpression.ReferencedDeclaration;
            if (indexedDeclaration == null)
            {
                return null;
            }

            if (indexedDeclaration.IsArray)
            {
                return ResolveLExpressionDeclaredTypeIsArray(dictionaryAccessExpression);
            }

            var asTypeName = indexedDeclaration.AsTypeName;
            var asTypeDeclaration = indexedDeclaration.AsTypeDeclaration;

            return ResolveDefaultMember(dictionaryAccessExpression, asTypeName, asTypeDeclaration, defaultMemberResolutionRecursionDepth);
        }

        private IBoundExpression ResolveDefaultMember(IBoundExpression lExpression, string asTypeName, Declaration asTypeDeclaration, int defaultMemberResolutionRecursionDepth)
        {
            /*
                The declared type of <l-expression> is Object or Variant, and <argument-list> contains no 
                named arguments. In this case, the index expression is classified as an unbound member with 
                a declared type of Variant, referencing <l-expression> with no member name. 
             */
            if (
                (Tokens.Variant.Equals(asTypeName, StringComparison.InvariantCultureIgnoreCase)
                    || Tokens.Object.Equals(asTypeName, StringComparison.InvariantCultureIgnoreCase))
                && !_argumentList.HasNamedArguments)
            {
                return new IndexExpression(null, ExpressionClassification.Unbound, _expression, lExpression, _argumentList, isDefaultMemberAccess: true);
            }
            /*
                The declared type of <l-expression> is a specific class, which has a public default Property 
                Get, Property Let, function or subroutine, and one of the following is true:
            */
            if (asTypeDeclaration is ClassModuleDeclaration classModule
                && classModule.DefaultMember is Declaration defaultMember
                && IsPropertyGetLetFunctionProcedure(defaultMember)
                && IsPublic(defaultMember))
            {
                var defaultMemberClassification = DefaultMemberExpressionClassification(defaultMember);

                /*
                    This default member’s parameter list is compatible with <argument-list>. In this case, the 
                    index expression references this default member and takes on its classification and 
                    declared type.  

                    TODO: Improve argument compatibility check.
                 */
                var parameters = ((IParameterizedDeclaration) defaultMember).Parameters.ToList();
                if (ArgumentListIsCompatible(parameters, _argumentList))
                {
                    return new IndexExpression(defaultMember, defaultMemberClassification, _expression, _lExpression, _argumentList, isDefaultMemberAccess: true);
                }

                /**
                    This default member can accept no parameters. In this case, the static analysis restarts 
                    recursively, as if this default member was specified instead for <l-expression> with the 
                    same <argument-list>.
                */
                if (parameters.Count(parameter => !parameter.IsOptional) == 0
                    && DEFAULT_MEMBER_RECURSION_LIMIT > defaultMemberResolutionRecursionDepth)
                {
                    return ResolveRecursiveDefaultMember(defaultMember, defaultMemberClassification, defaultMemberResolutionRecursionDepth);
                }
            }

            return null;
        }

        private static bool ArgumentListIsCompatible(ICollection<ParameterDeclaration> parameters, ArgumentList argumentList)
        {
            return (parameters.Count >= argumentList.Arguments.Count 
                        || parameters.Any(parameter => parameter.IsParamArray))
                    && parameters.Count(parameter => !parameter.IsOptional) <= argumentList.Arguments.Count;
        }

        private IBoundExpression ResolveRecursiveDefaultMember(Declaration defaultMember, ExpressionClassification defaultMemberClassification, int defaultMemberResolutionRecursionDepth)
        {
            var defaultMemberAsLExpression = new SimpleNameExpression(defaultMember, defaultMemberClassification, _expression);
            return Resolve(defaultMemberAsLExpression, defaultMemberResolutionRecursionDepth + 1);
        }

        private static ExpressionClassification DefaultMemberExpressionClassification(Declaration defaultMember)
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

        private IBoundExpression ResolveLExpressionDeclaredTypeIsArray(IBoundExpression lExpression)
        {
            var indexedDeclaration = lExpression.ReferencedDeclaration;
            if (!indexedDeclaration?.IsArray ?? false)
            {
                return null;
            }

            /*
                 The declared type of <l-expression> is an array type, an empty argument list has not already 
                 been specified for it, and one of the following is true:  
             */

            if (!_argumentList.HasArguments)
            {
                /*
                    <argument-list> represents an empty argument list. In this case, the index expression 
                    takes on the classification and declared type of <l-expression> and references the same 
                    array.  
                 */

                return new IndexExpression(indexedDeclaration, lExpression.Classification, _expression, lExpression, _argumentList);
            }

            if (!_argumentList.HasNamedArguments)
            {
                /*
                    <argument-list> represents an argument list with a number of positional arguments equal 
                    to the rank of the array, and with no named arguments. In this case, the index expression 
                    references an individual element of the array, is classified as a variable and has the 
                    declared type of the array’s element type.  

                    TODO: Implement compatibility checking / amend the grammar
                 */

                return new IndexExpression(indexedDeclaration, ExpressionClassification.Variable, _expression, lExpression, _argumentList, isArrayAccess: true);
            }

            return null;
        }

        private IBoundExpression ResolveLExpressionIsPropertyFunctionSubroutine(IBoundExpression lExpression)
        {
            /*
                    <l-expression> is classified as a property or function and its parameter list is compatible with 
                    <argument-list>. In this case, the index expression references <l-expression> and takes on its 
                    classification and declared type. 

                    <l-expression> is classified as a subroutine and its parameter list is compatible with <argument-
                    list>. In this case, the index expression references <l-expression> and takes on its classification 
                    and declared type.   

                    Note: We assume compatibility through enforcement by the VBE.
             */
             return new IndexExpression(lExpression.ReferencedDeclaration, lExpression.Classification, _expression, lExpression, _argumentList);
        }

        private IBoundExpression ResolveLExpressionIsUnbound(IBoundExpression lExpression)
        {
            /*
                 <l-expression> is classified as an unbound member. In this case, the index expression references 
                 <l-expression>, is classified as an unbound member and its declared type is Variant.  
            */
            return new IndexExpression(lExpression.ReferencedDeclaration, ExpressionClassification.Unbound, _expression, lExpression, _argumentList);
        }
    }
}
