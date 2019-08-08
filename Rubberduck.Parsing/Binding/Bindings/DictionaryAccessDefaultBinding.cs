using System;
using Rubberduck.Parsing.Symbols;
using System.Linq;
using Antlr4.Runtime;
using Antlr4.Runtime.Atn;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA.DeclarationCaching;

namespace Rubberduck.Parsing.Binding
{
    public sealed class DictionaryAccessDefaultBinding : IExpressionBinding
    {
        private readonly ParserRuleContext _expression;
        private readonly IExpressionBinding _lExpressionBinding;
        private IBoundExpression _lExpression;
        private readonly ArgumentList _argumentList;

        private const int DEFAULT_MEMBER_RECURSION_LIMIT = 32;
        private int _defaultMemberRecursionLimitCounter = 0;

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

            if (_lExpression.Classification == ExpressionClassification.ResolutionFailed)
            {
                ResolveArgumentList(null);
                return CreateFailedExpression(_lExpression);
            }

            var lDeclaration = _lExpression.ReferencedDeclaration;

            if (_lExpression.Classification == ExpressionClassification.Unbound)
            {
                /*
                     <l-expression> is classified as an unbound member. In this case, the dictionary access expression  
                    is classified as an unbound member with a declared type of Variant, referencing <l-expression> with no member name.
                */
                ResolveArgumentList(lDeclaration);
                return new DictionaryAccessExpression(null, ExpressionClassification.Unbound, _expression, _lExpression, _argumentList);
            }

            if (lDeclaration == null)
            {
                ResolveArgumentList(null);
                return CreateFailedExpression(_lExpression);
            }

            var asTypeName = lDeclaration.AsTypeName;
            var asTypeDeclaration = lDeclaration.AsTypeDeclaration;

            return ResolveViaDefaultMember(_lExpression, asTypeName, asTypeDeclaration);
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

        private IBoundExpression ResolveViaDefaultMember(IBoundExpression lExpression, string asTypeName, Declaration asTypeDeclaration)
        {
            if (Tokens.Variant.Equals(asTypeName, StringComparison.InvariantCultureIgnoreCase) 
                    || Tokens.Object.Equals(asTypeName, StringComparison.InvariantCultureIgnoreCase))
            {            
                /*
                    The declared type of <l-expression> is Object or Variant. 
                    In this case, the dictionary access expression is classified as an unbound member with 
                    a declared type of Variant, referencing <l-expression> with no member name. 
                */
                ResolveArgumentList(null);
                return new DictionaryAccessExpression(null, ExpressionClassification.Unbound, _expression, lExpression, _argumentList);
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
                ResolveArgumentList(null);
                return CreateFailedExpression(lExpression);
            }

            var defaultMemberClassification = DefaultMemberClassification(defaultMember);

            var parameters = ((IParameterizedDeclaration) defaultMember).Parameters.ToList();

            if (parameters.Count > 0 
                && parameters.Count(param => !param.IsOptional) <= 1 
                && Tokens.String.Equals(parameters[0].AsTypeName, StringComparison.InvariantCultureIgnoreCase))
            {
                /*
                    This default member’s parameter list is compatible with <argument-list>. In this case, the 
                    dictionary access expression references this default member and takes on its classification and 
                    declared type.  
                */
                ResolveArgumentList(defaultMember);
                return new DictionaryAccessExpression(defaultMember, defaultMemberClassification, _expression,
                    lExpression, _argumentList);
            }

            if (parameters.Count(param => !param.IsOptional) == 0 
                && DEFAULT_MEMBER_RECURSION_LIMIT > _defaultMemberRecursionLimitCounter)
            {
                /*
                    This default member cannot accept any parameters. In this case, the static analysis restarts 
                    recursively, as if this default member was specified instead for <l-expression> with the 
                    same <argument-list>.
                */
                _defaultMemberRecursionLimitCounter++;

                var defaultMemberAsLExpression =
                    new SimpleNameExpression(defaultMember, defaultMemberClassification, _expression);
                var defaultMemberAsTypeName = defaultMember.AsTypeName;
                var defaultMemberAsTypeDeclaration = defaultMember.AsTypeDeclaration;

                return ResolveViaDefaultMember(defaultMemberAsLExpression, defaultMemberAsTypeName,
                    defaultMemberAsTypeDeclaration);
            }

            ResolveArgumentList(null);
            return CreateFailedExpression(lExpression);
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

        private ExpressionClassification DefaultMemberClassification(Declaration defaultMember)
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