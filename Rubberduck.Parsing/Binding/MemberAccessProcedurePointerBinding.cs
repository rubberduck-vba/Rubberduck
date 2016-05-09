using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public sealed class MemberAccessProcedurePointerBinding : IExpressionBinding
    {
        private readonly DeclarationFinder _declarationFinder;
        private readonly Declaration _project;
        private readonly Declaration _module;
        private readonly Declaration _parent;
        private readonly VBAExpressionParser.MemberAccessExpressionContext _memberAccessExpression;
        private readonly VBAExpressionParser.MemberAccessExprContext _memberAccessExpr;
        private ParserRuleContext _unrestrictedNameContext;
        private readonly IExpressionBinding _lExpressionBinding;

        public MemberAccessProcedurePointerBinding(
            DeclarationFinder declarationFinder,
            Declaration project,
            Declaration module,
            Declaration parent,
            VBAExpressionParser.MemberAccessExpressionContext expression,
            ParserRuleContext unrestrictedNameContext,
            IExpressionBinding lExpressionBinding)
        {
            _declarationFinder = declarationFinder;
            _project = project;
            _module = module;
            _parent = parent;
            _memberAccessExpression = expression;
            _lExpressionBinding = lExpressionBinding;
            _unrestrictedNameContext = unrestrictedNameContext;
        }

        public MemberAccessProcedurePointerBinding(
            DeclarationFinder declarationFinder,
            Declaration project,
            Declaration module,
            Declaration parent,
            VBAExpressionParser.MemberAccessExprContext expression,
            ParserRuleContext unrestrictedNameContext,
            IExpressionBinding lExpressionBinding)
        {
            _declarationFinder = declarationFinder;
            _project = project;
            _module = module;
            _parent = parent;
            _memberAccessExpr = expression;
            _lExpressionBinding = lExpressionBinding;
            _unrestrictedNameContext = unrestrictedNameContext;
        }

        private ParserRuleContext GetExpressionContext()
        {
            if (_memberAccessExpression != null)
            {
                return _memberAccessExpression;
            }
            return _memberAccessExpr;
        }

        private string GetUnrestrictedName()
        {
            if (_memberAccessExpression != null)
            {
                return ExpressionName.GetName(_memberAccessExpression.unrestrictedName());
            }
            return ExpressionName.GetName(_memberAccessExpr.unrestrictedName());
        }

        public IBoundExpression Resolve()
        {
            IBoundExpression boundExpression = null;
            var lExpression = _lExpressionBinding.Resolve();
            if (lExpression == null)
            {
                return null;
            }
            string name = GetUnrestrictedName();
            if (lExpression.Classification != ExpressionClassification.ProceduralModule)
            {
                return null;
            }
            boundExpression = ResolveMemberInModule(lExpression, name, lExpression.ReferencedDeclaration, DeclarationType.Function, ExpressionClassification.Function);
            if (boundExpression != null)
            {
                return boundExpression;
            }
            boundExpression = ResolveMemberInModule(lExpression, name, lExpression.ReferencedDeclaration, DeclarationType.Procedure, ExpressionClassification.Subroutine);
            if (boundExpression != null)
            {
                return boundExpression;
            }
            return boundExpression;
        }

        private IBoundExpression ResolveMemberInModule(IBoundExpression lExpression, string name, Declaration module, DeclarationType memberType, ExpressionClassification classification)
        {
            /*
                A member access expression under the procedure pointer binding context is valid only if <l-
                expression> is classified as a procedural module, this procedural module has an accessible function 
                or subroutine with the same name value as <unrestricted-name>, and <unrestricted-name> either 
                does not specify a type character or specifies a type character whose associated type matches the 
                declared type of the function or subroutine. In this case, the member access expression is classified 
                as a function or subroutine, respectively.  
             */
            // AddressOf is only allowed in the same project. See The "procedure pointer binding context" for "simple name expressions" section in the MS-VBAL document.
            var enclosingProjectType = _declarationFinder.FindMemberEnclosedProjectInModule(_project, _module, _parent, module, name, memberType);
            if (enclosingProjectType != null)
            {
                return new MemberAccessExpression(enclosingProjectType, classification, GetExpressionContext(), _unrestrictedNameContext, lExpression);
            }
            return null;
        }
    }
}
