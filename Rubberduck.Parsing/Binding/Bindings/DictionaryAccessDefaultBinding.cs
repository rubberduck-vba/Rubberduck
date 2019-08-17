using System;
using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.Binding
{
    public sealed class DictionaryAccessDefaultBinding : IExpressionBinding
    {
        private readonly ParserRuleContext _expression;
        private readonly IExpressionBinding _lExpressionBinding;
        private IBoundExpression _lExpression;
        private readonly ArgumentList _argumentList;

        private const int DEFAULT_MEMBER_RECURSION_LIMIT = 32;

        //This is based on the spec at https://docs.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/MS-VBAL/f20c9ebc-3365-4614-9788-1cd50a504574

        public DictionaryAccessDefaultBinding(
            ParserRuleContext expression,
            IExpressionBinding lExpressionBinding,
            ArgumentList argumentList)
            : this(
                expression,
                (IBoundExpression) null,
                argumentList)
        {
            _lExpressionBinding = lExpressionBinding;
        }

        public DictionaryAccessDefaultBinding(
            ParserRuleContext expression,
            IBoundExpression lExpression,
            ArgumentList argumentList)
        {
            _expression = expression;
            _lExpression = lExpression;
            _argumentList = argumentList;
        }

        private static void ResolveArgumentList(Declaration calledProcedure, ArgumentList argumentList)
        {
            foreach (var argument in argumentList.Arguments)
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

            return Resolve(_lExpression, _argumentList, _expression);
        }

        private static IBoundExpression Resolve(IBoundExpression lExpression, ArgumentList argumentList, ParserRuleContext expression)
        {
            if (lExpression.Classification == ExpressionClassification.ResolutionFailed)
            {
                ResolveArgumentList(null, argumentList);
                return CreateFailedExpression(lExpression, argumentList);
            }

            var lDeclaration = lExpression.ReferencedDeclaration;

            if (lExpression.Classification == ExpressionClassification.Unbound)
            {
                /*
                     <l-expression> is classified as an unbound member. In this case, the dictionary access expression  
                    is classified as an unbound member with a declared type of Variant, referencing <l-expression> with no member name.
                */
                ResolveArgumentList(lDeclaration, argumentList);
                return new DictionaryAccessExpression(null, ExpressionClassification.Unbound, expression, lExpression, argumentList, 1);
            }

            if (lDeclaration == null)
            {
                ResolveArgumentList(null, argumentList);
                return CreateFailedExpression(lExpression, argumentList);
            }

            var asTypeName = lDeclaration.AsTypeName;
            var asTypeDeclaration = lDeclaration.AsTypeDeclaration;

            return ResolveViaDefaultMember(lExpression, asTypeName, asTypeDeclaration, argumentList, expression);
        }

        private static IBoundExpression CreateFailedExpression(IBoundExpression lExpression, ArgumentList argumentList)
        {
            var failedExpr = new ResolutionFailedExpression();
            failedExpr.AddSuccessfullyResolvedExpression(lExpression);
            foreach (var arg in argumentList.Arguments)
            {
                failedExpr.AddSuccessfullyResolvedExpression(arg.Expression);
            }

            return failedExpr;
        }

        private static IBoundExpression ResolveViaDefaultMember(IBoundExpression lExpression, string asTypeName, Declaration asTypeDeclaration, ArgumentList argumentList, ParserRuleContext expression, int recursionDepth = 1)
        {
            if (Tokens.Variant.Equals(asTypeName, StringComparison.InvariantCultureIgnoreCase) 
                    || Tokens.Object.Equals(asTypeName, StringComparison.InvariantCultureIgnoreCase))
            {            
                /*
                    The declared type of <l-expression> is Object or Variant. 
                    In this case, the dictionary access expression is classified as an unbound member with 
                    a declared type of Variant, referencing <l-expression> with no member name. 
                */
                ResolveArgumentList(null, argumentList);
                return new DictionaryAccessExpression(null, ExpressionClassification.Unbound, expression, lExpression, argumentList, recursionDepth);
            }

            /*
                The declared type of <l-expression> is a specific class, which has a public default Property 
                Get, Property Let, function or subroutine.
            */
            var defaultMember = (asTypeDeclaration as ClassModuleDeclaration)?.DefaultMember;
            if (defaultMember == null
                || !IsPropertyGetLetFunctionProcedure(defaultMember)
                || !IsPublic(defaultMember))
            {
                ResolveArgumentList(null, argumentList);
                return CreateFailedExpression(lExpression, argumentList);
            }

            var defaultMemberClassification = DefaultMemberClassification(defaultMember);

            var parameters = ((IParameterizedDeclaration) defaultMember).Parameters.ToList();

            if (IsCompatibleWithOneStringArgument(parameters))
            {
                /*
                    This default member’s parameter list is compatible with <argument-list>. In this case, the 
                    dictionary access expression references this default member and takes on its classification and 
                    declared type.  
                */
                ResolveArgumentList(defaultMember, argumentList);
                return new DictionaryAccessExpression(defaultMember, defaultMemberClassification, expression, lExpression, argumentList, recursionDepth);
            }

            if (parameters.Count(param => !param.IsOptional) == 0 
                && DEFAULT_MEMBER_RECURSION_LIMIT >= recursionDepth)
            {
                /*
                    This default member cannot accept any parameters. In this case, the static analysis restarts 
                    recursively, as if this default member was specified instead for <l-expression> with the 
                    same <argument-list>.
                */

                //In contrast to the IndexDefaultBinding we pass the original expression context since the default member accesses will be attached to the exclamation mark.
                return ResolveRecursiveDefaultMember(defaultMember, defaultMemberClassification, argumentList, expression, recursionDepth);
            }

            ResolveArgumentList(null, argumentList);
            return CreateFailedExpression(lExpression, argumentList);
        }

        private static bool IsCompatibleWithOneStringArgument(List<ParameterDeclaration> parameters)
        {
            return parameters.Count > 0 
                   && parameters.Count(param => !param.IsOptional) <= 1 
                   && (Tokens.String.Equals(parameters[0].AsTypeName, StringComparison.InvariantCultureIgnoreCase)
                       || Tokens.Variant.Equals(parameters[0].AsTypeName, StringComparison.InvariantCultureIgnoreCase));
        }

        private static IBoundExpression ResolveRecursiveDefaultMember(Declaration defaultMember, ExpressionClassification defaultMemberClassification, ArgumentList argumentList, ParserRuleContext expression, int recursionDepth)
        {
            var defaultMemberAsLExpression = new SimpleNameExpression(defaultMember, defaultMemberClassification, expression);
            var defaultMemberAsTypeName = defaultMember.AsTypeName;
            var defaultMemberAsTypeDeclaration = defaultMember.AsTypeDeclaration;

            return ResolveViaDefaultMember(
                defaultMemberAsLExpression,
                defaultMemberAsTypeName,
                defaultMemberAsTypeDeclaration,
                argumentList,
                expression,
                recursionDepth + 1);
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