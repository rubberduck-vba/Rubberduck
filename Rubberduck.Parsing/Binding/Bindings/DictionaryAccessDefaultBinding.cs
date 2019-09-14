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
            var arguments = argumentList.Arguments;
            for (var index = 0; index < arguments.Count; index++)
            {
                arguments[index].Resolve(calledProcedure, index);
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
                var argumentExpressions = argumentList.Arguments.Select(arg => arg.Expression);
                return lExpression.JoinAsFailedResolution(expression, argumentExpressions);
            }

            if (!(expression is VBAParser.LExpressionContext lExpressionContext))
            {
                ResolveArgumentList(null, argumentList);
                return CreateFailedExpression(lExpression, argumentList, expression);
            }

            var lDeclaration = lExpression.ReferencedDeclaration;
            var defaultMemberContext = DefaultMemberReferenceContext(lExpressionContext);

            if (lExpression.Classification == ExpressionClassification.Unbound)
            {
                /*
                     <l-expression> is classified as an unbound member. In this case, the dictionary access expression  
                    is classified as an unbound member with a declared type of Variant, referencing <l-expression> with no member name.
                */
                ResolveArgumentList(lDeclaration, argumentList);
                return new DictionaryAccessExpression(null, ExpressionClassification.Unbound, expression, lExpression, argumentList, defaultMemberContext,1, null);
            }

            if (lDeclaration == null)
            {
                ResolveArgumentList(null, argumentList);
                return CreateFailedExpression(lExpression, argumentList, expression);
            }

            var asTypeName = lDeclaration.AsTypeName;
            var asTypeDeclaration = lDeclaration.AsTypeDeclaration;

            return ResolveViaDefaultMember(lExpression, asTypeName, asTypeDeclaration, argumentList, lExpressionContext, defaultMemberContext);
        }

        private static IBoundExpression CreateFailedExpression(IBoundExpression lExpression, ArgumentList argumentList, ParserRuleContext context)
        {
            var failedExpr = new ResolutionFailedExpression(context, true);
            failedExpr.AddSuccessfullyResolvedExpression(lExpression);

            var argumentExpressions = argumentList.Arguments.Select(arg => arg.Expression);
            return failedExpr.JoinAsFailedResolution(context, argumentExpressions);
        }

        private static IBoundExpression CreateFailedDefaultMemberAccessExpression(IBoundExpression lExpression, ArgumentList argumentList, ParserRuleContext context)
        {
            var failedExpr = new DictionaryAccessExpression(lExpression.ReferencedDeclaration, ExpressionClassification.ResolutionFailed, context, lExpression, argumentList, context);

            var argumentExpressions = argumentList.Arguments.Select(arg => arg.Expression);
            return failedExpr.JoinAsFailedResolution(context, argumentExpressions.Concat(new []{ lExpression }));
        }

        private static IBoundExpression ResolveViaDefaultMember(IBoundExpression lExpression, string asTypeName, Declaration asTypeDeclaration, ArgumentList argumentList, ParserRuleContext expression, ParserRuleContext defaultMemberContext, int recursionDepth = 1, RecursiveDefaultMemberAccessExpression containedExpression = null)
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
                return new DictionaryAccessExpression(null, ExpressionClassification.Unbound, expression, lExpression, argumentList, defaultMemberContext, recursionDepth, containedExpression);
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
                return CreateFailedDefaultMemberAccessExpression(lExpression, argumentList, expression);
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
                return new DictionaryAccessExpression(defaultMember, ExpressionClassification.Variable, expression, lExpression, argumentList, defaultMemberContext, recursionDepth, containedExpression);
            }

            if (parameters.All(param => param.IsOptional) 
                && DEFAULT_MEMBER_RECURSION_LIMIT >= recursionDepth)
            {
                /*
                    This default member cannot accept any parameters. In this case, the static analysis restarts 
                    recursively, as if this default member was specified instead for <l-expression> with the 
                    same <argument-list>.
                */
                return ResolveRecursiveDefaultMember(lExpression, defaultMember, defaultMemberClassification, argumentList, expression, defaultMemberContext, recursionDepth, containedExpression);
            }

            ResolveArgumentList(null, argumentList);
            return CreateFailedDefaultMemberAccessExpression(lExpression, argumentList, expression);
        }

        private static bool IsCompatibleWithOneStringArgument(List<ParameterDeclaration> parameters)
        {
            return parameters.Count > 0 
                   && parameters.Count(param => !param.IsOptional && !param.IsParamArray) <= 1 
                   && (Tokens.String.Equals(parameters[0].AsTypeName, StringComparison.InvariantCultureIgnoreCase)
                       || Tokens.Variant.Equals(parameters[0].AsTypeName, StringComparison.InvariantCultureIgnoreCase));
        }

        private static IBoundExpression ResolveRecursiveDefaultMember(IBoundExpression lExpression, Declaration defaultMember, ExpressionClassification defaultMemberClassification, ArgumentList argumentList, ParserRuleContext expression, ParserRuleContext defaultMemberContext, int recursionDepth, RecursiveDefaultMemberAccessExpression containedExpression)
        {
            var defaultMemberAsTypeName = defaultMember.AsTypeName;
            var defaultMemberAsTypeDeclaration = defaultMember.AsTypeDeclaration;

            var defaultMemberExpression = new RecursiveDefaultMemberAccessExpression(defaultMember, defaultMemberClassification, defaultMemberContext, recursionDepth, containedExpression);

            return ResolveViaDefaultMember(
                lExpression,
                defaultMemberAsTypeName,
                defaultMemberAsTypeDeclaration,
                argumentList,
                expression,
                defaultMemberContext,
                recursionDepth + 1,
                defaultMemberExpression);
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

        private static ParserRuleContext DefaultMemberReferenceContext(VBAParser.LExpressionContext context)
        {
            if (context is VBAParser.DictionaryAccessExprContext dictionaryAccess)
            {
                return dictionaryAccess.dictionaryAccess();
            }

            return ((VBAParser.WithDictionaryAccessExprContext)context).dictionaryAccess();
        }
    }
}