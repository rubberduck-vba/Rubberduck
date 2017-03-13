using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;
using System.Linq;

namespace Rubberduck.Parsing.Binding
{
    public sealed class IndexDefaultBinding : IExpressionBinding
    {
        private readonly DeclarationFinder _declarationFinder;
        private readonly Declaration _project;
        private readonly Declaration _module;
        private readonly Declaration _parent;
        private readonly ParserRuleContext _expression;
        private readonly IExpressionBinding _lExpressionBinding;
        private IBoundExpression _lExpression;
        private readonly ArgumentList _argumentList;

        private const int DEFAULT_MEMBER_RECURSION_LIMIT = 32;
        private int _defaultMemberRecursionLimitCounter = 0;

        public IndexDefaultBinding(
            DeclarationFinder declarationFinder,
            Declaration project,
            Declaration module,
            Declaration parent,
            ParserRuleContext expression,
            IExpressionBinding lExpressionBinding,
            ArgumentList argumentList)
            : this(
                  declarationFinder,
                  project,
                  module,
                  parent,
                  expression,
                  (IBoundExpression)null,
                  argumentList)
        {
            _lExpressionBinding = lExpressionBinding;
        }

        public IndexDefaultBinding(
            DeclarationFinder declarationFinder,
            Declaration project,
            Declaration module,
            Declaration parent,
            ParserRuleContext expression,
            IBoundExpression lExpression,
            ArgumentList argumentList)
        {
            _declarationFinder = declarationFinder;
            _project = project;
            _module = module;
            _parent = parent;
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

        private IBoundExpression Resolve(IBoundExpression lExpression)
        {
            IBoundExpression boundExpression = null;
            if (lExpression.Classification == ExpressionClassification.ResolutionFailed)
            {
                return CreateFailedExpression(lExpression);
            }
            boundExpression = ResolveLExpressionIsVariablePropertyFunctionNoParameters(lExpression);
            if (boundExpression != null)
            {
                return boundExpression;
            }
            boundExpression = ResolveLExpressionIsPropertyFunctionSubroutine(lExpression);
            if (boundExpression != null)
            {
                return boundExpression;
            }
            boundExpression = ResolveLExpressionIsUnbound(lExpression);
            if (boundExpression != null)
            {
                return boundExpression;
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

        private IBoundExpression ResolveLExpressionIsVariablePropertyFunctionNoParameters(IBoundExpression lExpression)
        {
            /*
             <l-expression> is classified as a variable, or <l-expression> is classified as a property or function 
                    with a parameter list that cannot accept any parameters and an <argument-list> that is not 
                    empty, and one of the following is true (see below):
             */
            bool isVariable = lExpression.Classification == ExpressionClassification.Variable;
            bool propertyWithParameters = lExpression.Classification == ExpressionClassification.Property && !((IParameterizedDeclaration)lExpression.ReferencedDeclaration).Parameters.Any();
            bool functionWithParameters = lExpression.Classification == ExpressionClassification.Function && !((IParameterizedDeclaration)lExpression.ReferencedDeclaration).Parameters.Any();
            if (isVariable || ((propertyWithParameters || functionWithParameters) && _argumentList.HasArguments))
            {
                IBoundExpression boundExpression = null;
                var asTypeName = lExpression.ReferencedDeclaration.AsTypeName;
                var asTypeDeclaration = lExpression.ReferencedDeclaration.AsTypeDeclaration;
                boundExpression = ResolveDefaultMember(lExpression, asTypeName, asTypeDeclaration);
                if (boundExpression != null)
                {
                    return boundExpression;
                }
                boundExpression = ResolveLExpressionDeclaredTypeIsArray(lExpression, asTypeDeclaration);
                if (boundExpression != null)
                {
                    return boundExpression;
                }
                return boundExpression;
            }
            return null;
        }

        private IBoundExpression ResolveDefaultMember(IBoundExpression lExpression, string asTypeName, Declaration asTypeDeclaration)
        {
            if (lExpression.ReferencedDeclaration.IsArray)
            {
                return null;
            }
            /*
                The declared type of <l-expression> is Object or Variant, and <argument-list> contains no 
                named arguments. In this case, the index expression is classified as an unbound member with 
                a declared type of Variant, referencing <l-expression> with no member name. 
             */
            if (
                asTypeName != null
                && (asTypeName.ToUpperInvariant() == "VARIANT" || asTypeName.ToUpperInvariant() == "OBJECT")
                && !_argumentList.HasNamedArguments)
            {
                return new IndexExpression(null, ExpressionClassification.Unbound, _expression, lExpression, _argumentList);
            }
            /*
                The declared type of <l-expression> is a specific class, which has a public default Property 
                Get, Property Let, function or subroutine, and one of the following is true:
            */
            bool hasDefaultMember = asTypeDeclaration != null
                && asTypeDeclaration.DeclarationType == DeclarationType.ClassModule
                && ((ClassModuleDeclaration)asTypeDeclaration).DefaultMember != null;
            if (hasDefaultMember)
            {
                ClassModuleDeclaration classModule = (ClassModuleDeclaration)asTypeDeclaration;
                Declaration defaultMember = classModule.DefaultMember;
                bool isPropertyGetLetFunctionProcedure =
                    defaultMember.DeclarationType == DeclarationType.PropertyGet
                    || defaultMember.DeclarationType == DeclarationType.PropertyLet
                    || defaultMember.DeclarationType == DeclarationType.Function
                    || defaultMember.DeclarationType == DeclarationType.Procedure;
                bool isPublic =
                    defaultMember.Accessibility == Accessibility.Global
                    || defaultMember.Accessibility == Accessibility.Implicit
                    || defaultMember.Accessibility == Accessibility.Public;
                if (isPropertyGetLetFunctionProcedure && isPublic)
                {

                    /*
                        This default member’s parameter list is compatible with <argument-list>. In this case, the 
                        index expression references this default member and takes on its classification and 
                        declared type.  

                        TODO: Primitive argument compatibility checking for now.
                     */
                    if (((IParameterizedDeclaration)defaultMember).Parameters.Count() == _argumentList.Arguments.Count)
                    {
                        return new IndexExpression(defaultMember, lExpression.Classification, _expression, lExpression, _argumentList);
                    }

                    /**
                        This default member cannot accept any parameters. In this case, the static analysis restarts 
                        recursively, as if this default member was specified instead for <l-expression> with the 
                        same <argument-list>.
                    */
                    if (((IParameterizedDeclaration)defaultMember).Parameters.Count() == 0)
                    {
                        // Recursion limit reached, abort.
                        if (DEFAULT_MEMBER_RECURSION_LIMIT == _defaultMemberRecursionLimitCounter)
                        {
                            return null;
                        }
                        _defaultMemberRecursionLimitCounter++;
                        ExpressionClassification classification;
                        if (defaultMember.DeclarationType.HasFlag(DeclarationType.Property))
                        {
                            classification = ExpressionClassification.Property;
                        }
                        else if (defaultMember.DeclarationType == DeclarationType.Procedure)
                        {
                            classification = ExpressionClassification.Subroutine;
                        }
                        else
                        {
                            classification = ExpressionClassification.Function;
                        }
                        var defaultMemberAsLExpression = new SimpleNameExpression(defaultMember, classification, _expression);
                        return Resolve(defaultMemberAsLExpression);
                    }
                }
            }
            return null;
        }

        private IBoundExpression ResolveLExpressionDeclaredTypeIsArray(IBoundExpression lExpression, Declaration asTypeDeclaration)
        {
            /*
                 The declared type of <l-expression> is an array type, an empty argument list has not already 
                 been specified for it, and one of the following is true:  
             */
            if (lExpression.ReferencedDeclaration.IsArray)
            {
                /*
                    <argument-list> represents an empty argument list. In this case, the index expression 
                    takes on the classification and declared type of <l-expression> and references the same 
                    array.  
                 */
                if (!_argumentList.HasArguments)
                {
                    return new IndexExpression(asTypeDeclaration, lExpression.Classification, _expression, lExpression, _argumentList);
                }
                else
                {
                    /*
                        <argument-list> represents an argument list with a number of positional arguments equal 
                        to the rank of the array, and with no named arguments. In this case, the index expression 
                        references an individual element of the array, is classified as a variable and has the 
                        declared type of the array’s element type.  

                        TODO: Implement compatibility checking / amend the grammar
                     */
                    if (!_argumentList.HasNamedArguments)
                    {
                        return new IndexExpression(asTypeDeclaration, ExpressionClassification.Variable, _expression, lExpression, _argumentList);
                    }
                }
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
            if (lExpression.Classification == ExpressionClassification.Property
               || lExpression.Classification == ExpressionClassification.Function
               || lExpression.Classification == ExpressionClassification.Subroutine)
            {
                return new IndexExpression(lExpression.ReferencedDeclaration, lExpression.Classification, _expression, lExpression, _argumentList);
            }
            return null;
        }

        private IBoundExpression ResolveLExpressionIsUnbound(IBoundExpression lExpression)
        {
            /*
                 <l-expression> is classified as an unbound member. In this case, the index expression references 
                 <l-expression>, is classified as an unbound member and its declared type is Variant.  
            */
            if (lExpression.Classification == ExpressionClassification.Unbound)
            {
                return new IndexExpression(lExpression.ReferencedDeclaration, ExpressionClassification.Unbound, _expression, lExpression, _argumentList);
            }
            return null;
        }
    }
}
