using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public sealed class MemberAccessTypeBinding : IExpressionBinding
    {
        private readonly DeclarationFinder _declarationFinder;
        private readonly Declaration _project;
        private readonly Declaration _module;
        private readonly Declaration _parent;
        private readonly VBAParser.MemberAccessExprContext _expression;
        private ParserRuleContext _unrestrictedNameContext;
        private readonly IExpressionBinding _lExpressionBinding;

        public MemberAccessTypeBinding(
            DeclarationFinder declarationFinder,
            Declaration project,
            Declaration module,
            Declaration parent,
            VBAParser.MemberAccessExprContext expression,
            ParserRuleContext unrestrictedNameContext,
            IExpressionBinding lExpressionBinding)
        {
            _declarationFinder = declarationFinder;
            _project = project;
            _module = module;
            _parent = parent;
            _expression = expression;
            _lExpressionBinding = lExpressionBinding;
            _unrestrictedNameContext = unrestrictedNameContext;
        }

        public IExpressionBinding LExpressionBinding
        {
            get
            {
                return _lExpressionBinding;
            }
        }

        public IBoundExpression Resolve()
        {
            IBoundExpression boundExpression = null;
            var lExpression = _lExpressionBinding.Resolve();
            if (lExpression.Classification == ExpressionClassification.ResolutionFailed)
            {
                return lExpression;
            }
            string unrestrictedName = Identifier.GetName(_expression.unrestrictedIdentifier());
            boundExpression = ResolveLExpressionIsProject(lExpression, unrestrictedName);
            if (boundExpression != null)
            {
                return boundExpression;
            }
            boundExpression = ResolveLExpressionIsModule(lExpression, unrestrictedName);
            if (boundExpression != null)
            {
                return boundExpression;
            }
            var failedExpr = new ResolutionFailedExpression();
            failedExpr.AddSuccessfullyResolvedExpression(lExpression);
            return failedExpr;
        }

        private IBoundExpression ResolveLExpressionIsProject(IBoundExpression lExpression, string name)
        {
            if (lExpression.Classification != ExpressionClassification.Project)
            {
                return null;
            }
            IBoundExpression boundExpression = null;
            var referencedProject = lExpression.ReferencedDeclaration;
            bool lExpressionIsEnclosingProject = _project.Equals(referencedProject);
            boundExpression = ResolveProject(lExpression, name);
            if (boundExpression != null)
            {
                return boundExpression;
            }
            boundExpression = ResolveProceduralModule(lExpressionIsEnclosingProject, lExpression, name, referencedProject);
            if (boundExpression != null)
            {
                return boundExpression;
            }
            boundExpression = ResolveClassModule(lExpressionIsEnclosingProject, lExpression, name, referencedProject);
            if (boundExpression != null)
            {
                return boundExpression;
            }
            boundExpression = ResolveMemberInReferencedProject(lExpressionIsEnclosingProject, lExpression, name, referencedProject, DeclarationType.UserDefinedType);
            if (boundExpression != null)
            {
                return boundExpression;
            }
            boundExpression = ResolveMemberInReferencedProject(lExpressionIsEnclosingProject, lExpression, name, referencedProject, DeclarationType.Enumeration);
            if (boundExpression != null)
            {
                return boundExpression;
            }
            return boundExpression;
        }

        private IBoundExpression ResolveProject(IBoundExpression lExpression, string name)
        {
            /*
              <l-expression> refers to the enclosing project and <unrestricted-name> is either the name of 
                the enclosing project or a referenced project. In this case, the member access expression is 
                classified as a project and refers to the specified project. 
             */
            if (_declarationFinder.IsMatch(_project.ProjectName, name))
            {
                return new MemberAccessExpression(_project, ExpressionClassification.Project, _expression, _unrestrictedNameContext, lExpression);
            }
            var referencedProjectRightOfDot = _declarationFinder.FindReferencedProject(_project, name);
            if (referencedProjectRightOfDot != null)
            {
                return new MemberAccessExpression(referencedProjectRightOfDot, ExpressionClassification.Project, _expression, _unrestrictedNameContext, lExpression);
            }
            return null;
        }

        private IBoundExpression ResolveProceduralModule(bool lExpressionIsEnclosingProject, IBoundExpression lExpression, string name, Declaration referencedProject)
        {
            /*
                The project has an accessible procedural module named <unrestricted-name>. In this case, the 
                member access expression is classified as a procedural module and refers to the specified 
                procedural module.  
             */
            if (lExpressionIsEnclosingProject)
            {
                if (_module.DeclarationType == DeclarationType.ProceduralModule && _declarationFinder.IsMatch(_module.IdentifierName, name))
                {
                    return new MemberAccessExpression(_module, ExpressionClassification.ProceduralModule, _expression, _unrestrictedNameContext, lExpression);
                }
                var proceduralModuleEnclosingProject = _declarationFinder.FindModuleEnclosingProjectWithoutEnclosingModule(_project, _module, name, DeclarationType.ProceduralModule);
                if (proceduralModuleEnclosingProject != null)
                {
                    return new MemberAccessExpression(proceduralModuleEnclosingProject, ExpressionClassification.ProceduralModule, _expression, _unrestrictedNameContext, lExpression);
                }
            }
            else
            {
                var proceduralModuleInReferencedProject = _declarationFinder.FindModuleReferencedProject(_project, _module, referencedProject, name, DeclarationType.ProceduralModule);
                if (proceduralModuleInReferencedProject != null)
                {
                    return new MemberAccessExpression(proceduralModuleInReferencedProject, ExpressionClassification.ProceduralModule, _expression, _unrestrictedNameContext, lExpression);
                }
            }
            return null;
        }

        private IBoundExpression ResolveClassModule(bool lExpressionIsEnclosingProject, IBoundExpression lExpression, string name, Declaration referencedProject)
        {
            /*
                The project has an accessible class module named <unrestricted-name>. In this case, the 
                member access expression is classified as a type and refers to the specified class.  
             */
            if (lExpressionIsEnclosingProject)
            {
                if (_module.DeclarationType == DeclarationType.ClassModule && _declarationFinder.IsMatch(_module.IdentifierName, name))
                {
                    return new MemberAccessExpression(_module, ExpressionClassification.Type, _expression, _unrestrictedNameContext, lExpression);
                }
                var classModuleEnclosingProject = _declarationFinder.FindModuleEnclosingProjectWithoutEnclosingModule(_project, _module, name, DeclarationType.ClassModule);
                if (classModuleEnclosingProject != null)
                {
                    return new MemberAccessExpression(classModuleEnclosingProject, ExpressionClassification.Type, _expression, _unrestrictedNameContext, lExpression);
                }
            }
            else
            {
                var classModuleInReferencedProject = _declarationFinder.FindModuleReferencedProject(_project, _module, referencedProject, name, DeclarationType.ClassModule);
                if (classModuleInReferencedProject != null)
                {
                    return new MemberAccessExpression(classModuleInReferencedProject, ExpressionClassification.Type, _expression, _unrestrictedNameContext, lExpression);
                }
            }
            return null;
        }

        private IBoundExpression ResolveMemberInReferencedProject(bool lExpressionIsEnclosingProject, IBoundExpression lExpression, string name, Declaration referencedProject, DeclarationType memberType)
        {
            /*
                The project does not have an accessible module named <unrestricted-name> and exactly one of 
                the procedural modules within the project contains a UDT or Enum definition named 
                <unrestricted-name>. In this case, the member access expression is classified as a type and 
                refers to the specified UDT or enum. 
             */
            if (lExpressionIsEnclosingProject)
            {
                var foundType = _declarationFinder.FindMemberEnclosingModule(_module, _parent, name, memberType);
                if (foundType != null)
                {
                    return new MemberAccessExpression(foundType, ExpressionClassification.Type, _expression, _unrestrictedNameContext, lExpression);
                }
                var accessibleType = _declarationFinder.FindMemberEnclosedProjectWithoutEnclosingModule(_project, _module, _parent, name, memberType);
                if (accessibleType != null)
                {
                    return new MemberAccessExpression(accessibleType, ExpressionClassification.Type, _expression, _unrestrictedNameContext, lExpression);
                }
            }
            else
            {
                var referencedProjectType = _declarationFinder.FindMemberReferencedProject(_project, _module, _parent, referencedProject, name, memberType);
                if (referencedProjectType != null)
                {
                    return new MemberAccessExpression(referencedProjectType, ExpressionClassification.Type, _expression, _unrestrictedNameContext, lExpression);
                }
            }
            return null;
        }

        private IBoundExpression ResolveLExpressionIsModule(IBoundExpression lExpression, string name)
        {
            if (lExpression.Classification != ExpressionClassification.ProceduralModule && lExpression.Classification != ExpressionClassification.Type)
            {
                return null;
            }
            IBoundExpression boundExpression = null;
            boundExpression = ResolveMemberInModule(lExpression, name, lExpression.ReferencedDeclaration, DeclarationType.UserDefinedType);
            if (boundExpression != null)
            {
                return boundExpression;
            }
            boundExpression = ResolveMemberInModule(lExpression, name, lExpression.ReferencedDeclaration, DeclarationType.Enumeration);
            if (boundExpression != null)
            {
                return boundExpression;
            }
            return boundExpression;
        }

        private IBoundExpression ResolveMemberInModule(IBoundExpression lExpression, string name, Declaration module, DeclarationType memberType)
        {
            /*
                <l-expression> is classified as a procedural module or a type referencing a class defined in a 
                class module, and one of the following is true:  

                This module has an accessible UDT or Enum definition named <unrestricted-name>. In this 
                case, the member access expression is classified as a type and refers to the specified UDT or 
                Enum type.  
             */
            var enclosingProjectType = _declarationFinder.FindMemberEnclosedProjectInModule(_project, _module, _parent, module, name, memberType);
            if (enclosingProjectType != null)
            {
                return new MemberAccessExpression(enclosingProjectType, ExpressionClassification.Type, _expression, _unrestrictedNameContext, lExpression);
            }
       
            var referencedProjectType = _declarationFinder.FindMemberReferencedProjectInModule(_project, _module, _parent, module, name, memberType);
            if (referencedProjectType != null)
            {
                return new MemberAccessExpression(referencedProjectType, ExpressionClassification.Type, _expression, _unrestrictedNameContext, lExpression);
            }
            return null;
        }
    }
}
