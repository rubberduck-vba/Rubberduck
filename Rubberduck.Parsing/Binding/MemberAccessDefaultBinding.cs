using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public sealed class MemberAccessDefaultBinding : IExpressionBinding
    {
        private readonly DeclarationFinder _declarationFinder;
        private readonly Declaration _project;
        private readonly Declaration _module;
        private readonly Declaration _parent;
        private readonly VBAExpressionParser.MemberAccessExpressionContext _memberAccessExpression;
        private readonly VBAExpressionParser.MemberAccessExprContext _memberAccessExpr;
        private readonly IExpressionBinding _lExpressionBinding;

        public MemberAccessDefaultBinding(
            DeclarationFinder declarationFinder,
            Declaration project,
            Declaration module,
            Declaration parent,
            VBAExpressionParser.MemberAccessExpressionContext expression,
            IExpressionBinding lExpressionBinding)
        {
            _declarationFinder = declarationFinder;
            _project = project;
            _module = module;
            _parent = parent;
            _memberAccessExpression = expression;
            _lExpressionBinding = lExpressionBinding;
        }

        public MemberAccessDefaultBinding(
            DeclarationFinder declarationFinder,
            Declaration project,
            Declaration module,
            Declaration parent,
            VBAExpressionParser.MemberAccessExprContext expression,
            IExpressionBinding lExpressionBinding)
        {
            _declarationFinder = declarationFinder;
            _project = project;
            _module = module;
            _parent = parent;
            _memberAccessExpr = expression;
            _lExpressionBinding = lExpressionBinding;
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
            string unrestrictedName = GetUnrestrictedName();
            boundExpression = ResolveLExpressionIsVariablePropertyOrFunction(lExpression, unrestrictedName);
            if (boundExpression != null)
            {
                return boundExpression;
            }
            boundExpression = ResolveLExpressionIsUnbound(lExpression, unrestrictedName);
            if (boundExpression != null)
            {
                return boundExpression;
            }
            boundExpression = ResolveLExpressionIsProject(lExpression, unrestrictedName);
            if (boundExpression != null)
            {
                return boundExpression;
            }
            boundExpression = ResolveLExpressionIsProceduralModule(lExpression, unrestrictedName);
            if (boundExpression != null)
            {
                return boundExpression;
            }
            boundExpression = ResolveLExpressionIsEnum(lExpression, unrestrictedName);
            if (boundExpression != null)
            {
                return boundExpression;
            }
            return null;
        }

        private IBoundExpression ResolveLExpressionIsVariablePropertyOrFunction(IBoundExpression lExpression, string name)
        {
            /*
              
                <l-expression> is classified as a variable, a property or a function and one of the following is 
                true:  
                    1. The declared type of <l-expression> is a UDT type or specific class, this type has an accessible 
                    member named <unrestricted-name>, <unrestricted-name> either does not specify a type 
                    character or specifies a type character whose associated type matches the declared type of 
                    the member, and one of the following is true:

                        1.1 The member is a variable, property or function. In this case, the member access expression 
                        is classified as a variable, property or function, respectively, refers to the member, and has 
                        the same declared type as the member.

                        1.2 The member is a subroutine. In this case, the member access expression is classified as a 
                        subroutine and refers to the member.


                    2. The declared type of <l-expression> is Object or Variant. In this case, the member access 
                    expression is classified as an unbound member and has a declared type of Variant.  
             */
            if (
                lExpression.Classification != ExpressionClassification.Variable
                && lExpression.Classification != ExpressionClassification.Property
                && lExpression.Classification != ExpressionClassification.Function)
            {
                return null;
            }
            var lExpressionDeclaration = lExpression.ReferencedDeclaration;
            var referencedType = lExpressionDeclaration.AsTypeDeclaration;
            if (referencedType == null)
            {
                return null;
            }
            if (referencedType.DeclarationType != DeclarationType.UserDefinedType && referencedType.DeclarationType != DeclarationType.ClassModule)
            {
                return null;
            }
            var variable = _declarationFinder.FindMemberWithParent(_project, _module, referencedType, name, DeclarationType.Variable);
            if (variable != null)
            {
                return new MemberAccessExpression(variable, ExpressionClassification.Variable, GetExpressionContext(), lExpression);
            }
            var property = _declarationFinder.FindMemberWithParent(_project, _module, referencedType, name, DeclarationType.Property);
            if (property != null)
            {
                return new MemberAccessExpression(property, ExpressionClassification.Property, GetExpressionContext(), lExpression);
            }
            var function = _declarationFinder.FindMemberWithParent(_project, _module, referencedType, name, DeclarationType.Function);
            if (function != null)
            {
                return new MemberAccessExpression(function, ExpressionClassification.Function, GetExpressionContext(), lExpression);
            }
            var subroutine = _declarationFinder.FindMemberWithParent(_project, _module, referencedType, name, DeclarationType.Procedure);
            if (subroutine != null)
            {
                return new MemberAccessExpression(subroutine, ExpressionClassification.Subroutine, GetExpressionContext(), lExpression);
            }
            // Note: To not have to deal with declared types we simply assume that no match means unbound member.
            // This way the rest of the member access expression can still be bound.
            return new MemberAccessExpression(null, ExpressionClassification.Unbound, GetExpressionContext(), lExpression);
        }

        private IBoundExpression ResolveLExpressionIsUnbound(IBoundExpression lExpression, string name)
        {
            if (lExpression.Classification != ExpressionClassification.Unbound)
            {
                return null;
            }
            return new MemberAccessExpression(null, ExpressionClassification.Unbound, GetExpressionContext(), lExpression);
        }

        private IBoundExpression ResolveLExpressionIsProject(IBoundExpression lExpression, string name)
        {
            /*
                <l-expression> is classified as a project, this project is either the enclosing project or a 
                referenced project, and one of the following is true:  
                    -   <l-expression> refers to the enclosing project and <unrestricted-name> is either the name of 
                        the enclosing project or a referenced project. In this case, the member access expression is 
                        classified as a project and refers to the specified project.  
                    -   The project has an accessible procedural module named <unrestricted-name>. In this case, the 
                        member access expression is classified as a procedural module and refers to the specified 
                        procedural module.  
                    -   The project does not have an accessible procedural module named <unrestricted-name> and 
                        exactly one of the procedural modules within the project has an accessible member named 
                        <unrestricted-name>, <unrestricted-name> either does not specify a type character or 
                        specifies a type character whose associated type matches the declared type of the member, 
                        and one of the following is true:  
                        -   The member is a variable, property or function. In this case, the member access expression 
                            is classified as a variable, property or function, respectively, refers to the member, and has 
                            the same declared type as the member.  
                        -   The member is a subroutine. In this case, the member access expression is classified as a 
                            subroutine and refers to the member.  
                        -   The member is a value. In this case, the member access expression is classified as a value 
                            with the same declared type as the member.  
             */
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
            boundExpression = ResolveMemberInReferencedProject(lExpressionIsEnclosingProject, lExpression, name, referencedProject, DeclarationType.Variable, ExpressionClassification.Variable);
            if (boundExpression != null)
            {
                return boundExpression;
            }
            boundExpression = ResolveMemberInReferencedProject(lExpressionIsEnclosingProject, lExpression, name, referencedProject, DeclarationType.Property, ExpressionClassification.Property);
            if (boundExpression != null)
            {
                return boundExpression;
            }
            boundExpression = ResolveMemberInReferencedProject(lExpressionIsEnclosingProject, lExpression, name, referencedProject, DeclarationType.Function, ExpressionClassification.Function);
            if (boundExpression != null)
            {
                return boundExpression;
            }
            boundExpression = ResolveMemberInReferencedProject(lExpressionIsEnclosingProject, lExpression, name, referencedProject, DeclarationType.Procedure, ExpressionClassification.Subroutine);
            if (boundExpression != null)
            {
                return boundExpression;
            }
            boundExpression = ResolveMemberInReferencedProject(lExpressionIsEnclosingProject, lExpression, name, referencedProject, DeclarationType.Constant, ExpressionClassification.Value);
            if (boundExpression != null)
            {
                return boundExpression;
            }
            boundExpression = ResolveMemberInReferencedProject(lExpressionIsEnclosingProject, lExpression, name, referencedProject, DeclarationType.Enumeration, ExpressionClassification.Value);
            if (boundExpression != null)
            {
                return boundExpression;
            }
            boundExpression = ResolveMemberInReferencedProject(lExpressionIsEnclosingProject, lExpression, name, referencedProject, DeclarationType.EnumerationMember, ExpressionClassification.Value);
            if (boundExpression != null)
            {
                return boundExpression;
            }
            return boundExpression;
        }

        private IBoundExpression ResolveProject(IBoundExpression lExpression, string name)
        {
            if (_declarationFinder.IsMatch(_project.ProjectName, name))
            {
                return new MemberAccessExpression(_project, ExpressionClassification.Project, GetExpressionContext(), lExpression);
            }
            var referencedProjectRightOfDot = _declarationFinder.FindReferencedProject(_project, name);
            if (referencedProjectRightOfDot != null)
            {
                return new MemberAccessExpression(referencedProjectRightOfDot, ExpressionClassification.Project, GetExpressionContext(), lExpression);
            }
            return null;
        }

        private IBoundExpression ResolveProceduralModule(bool lExpressionIsEnclosingProject, IBoundExpression lExpression, string name, Declaration referencedProject)
        {
            if (lExpressionIsEnclosingProject)
            {
                if (_module.DeclarationType == DeclarationType.ProceduralModule && _declarationFinder.IsMatch(_module.IdentifierName, name))
                {
                    return new MemberAccessExpression(_module, ExpressionClassification.ProceduralModule, GetExpressionContext(), lExpression);
                }
                var proceduralModuleEnclosingProject = _declarationFinder.FindModuleEnclosingProjectWithoutEnclosingModule(_project, _module, name, DeclarationType.ProceduralModule);
                if (proceduralModuleEnclosingProject != null)
                {
                    return new MemberAccessExpression(proceduralModuleEnclosingProject, ExpressionClassification.ProceduralModule, GetExpressionContext(), lExpression);
                }
            }
            else
            {
                var proceduralModuleInReferencedProject = _declarationFinder.FindModuleReferencedProject(_project, _module, referencedProject, name, DeclarationType.ProceduralModule);
                if (proceduralModuleInReferencedProject != null)
                {
                    return new MemberAccessExpression(proceduralModuleInReferencedProject, ExpressionClassification.ProceduralModule, GetExpressionContext(), lExpression);
                }
            }
            return null;
        }

        private IBoundExpression ResolveMemberInReferencedProject(bool lExpressionIsEnclosingProject, IBoundExpression lExpression, string name, Declaration referencedProject, DeclarationType memberType, ExpressionClassification classification)
        {
            if (lExpressionIsEnclosingProject)
            {
                var foundType = _declarationFinder.FindMemberEnclosingModule(_project, _module, _parent, name, memberType);
                if (foundType != null)
                {
                    return new MemberAccessExpression(foundType, classification, GetExpressionContext(), lExpression);
                }
                var accessibleType = _declarationFinder.FindMemberEnclosedProjectWithoutEnclosingModule(_project, _module, _parent, name, memberType);
                if (accessibleType != null)
                {
                    return new MemberAccessExpression(accessibleType, classification, GetExpressionContext(), lExpression);
                }
            }
            else
            {
                var referencedProjectType = _declarationFinder.FindMemberReferencedProject(_project, _module, _parent, referencedProject, name, memberType);
                if (referencedProjectType != null)
                {
                    return new MemberAccessExpression(referencedProjectType, classification, GetExpressionContext(), lExpression);
                }
            }
            return null;
        }

        private IBoundExpression ResolveLExpressionIsProceduralModule(IBoundExpression lExpression, string name)
        {
            /*
                <l-expression> is classified as a procedural module, this procedural module has an accessible 
                member named <unrestricted-name>, <unrestricted-name> either does not specify a type 
                character or specifies a type character whose associated type matches the declared type of the 
                member, and one of the following is true:  
                    -   The member is a variable, property or function. In this case, the member access expression is 
                        classified as a variable, property or function, respectively, and has the same declared type as 
                        the member.  
                    -   The member is a subroutine. In this case, the member access expression is classified as a 
                        subroutine.  
                    -   The member is a value. In this case, the member access expression is classified as a value with 
                        the same declared type as the member. 
             */
            if (lExpression.Classification != ExpressionClassification.ProceduralModule)
            {
                return null;
            }
            IBoundExpression boundExpression = null;
            boundExpression = ResolveMemberInModule(lExpression, name, lExpression.ReferencedDeclaration, DeclarationType.Variable, ExpressionClassification.Variable);
            if (boundExpression != null)
            {
                return boundExpression;
            }
            boundExpression = ResolveMemberInModule(lExpression, name, lExpression.ReferencedDeclaration, DeclarationType.Property, ExpressionClassification.Property);
            if (boundExpression != null)
            {
                return boundExpression;
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
            boundExpression = ResolveMemberInModule(lExpression, name, lExpression.ReferencedDeclaration, DeclarationType.Constant, ExpressionClassification.Value);
            if (boundExpression != null)
            {
                return boundExpression;
            }
            boundExpression = ResolveMemberInModule(lExpression, name, lExpression.ReferencedDeclaration, DeclarationType.Enumeration, ExpressionClassification.Value);
            if (boundExpression != null)
            {
                return boundExpression;
            }
            boundExpression = ResolveMemberInModule(lExpression, name, lExpression.ReferencedDeclaration, DeclarationType.EnumerationMember, ExpressionClassification.Value);
            if (boundExpression != null)
            {
                return boundExpression;
            }
            return boundExpression;
        }

        private IBoundExpression ResolveMemberInModule(IBoundExpression lExpression, string name, Declaration module, DeclarationType memberType, ExpressionClassification classification)
        {
            var enclosingProjectType = _declarationFinder.FindMemberEnclosedProjectInModule(_project, _module, _parent, module, name, memberType);
            if (enclosingProjectType != null)
            {
                return new MemberAccessExpression(enclosingProjectType, classification, GetExpressionContext(), lExpression);
            }

            var referencedProjectType = _declarationFinder.FindMemberReferencedProjectInModule(_project, _module, _parent, module, name, memberType);
            if (referencedProjectType != null)
            {
                return new MemberAccessExpression(referencedProjectType, classification, GetExpressionContext(), lExpression);
            }
            return null;
        }

        private IBoundExpression ResolveLExpressionIsEnum(IBoundExpression lExpression, string name)
        {
            /*
                <l-expression> is classified as a type, this type is an Enum type, and this type has an enum 
                member named <unrestricted-name>. In this case, the member access expression is classified 
                as a value with the same declared type as the enum member.  
             */
            if (lExpression.Classification != ExpressionClassification.Type && lExpression.ReferencedDeclaration.DeclarationType != DeclarationType.Enumeration)
            {
                return null;
            }
            var enumMember = _declarationFinder.FindMemberWithParent(_project, _module, lExpression.ReferencedDeclaration, name, DeclarationType.EnumerationMember);
            if (enumMember != null)
            {
                return new MemberAccessExpression(enumMember, ExpressionClassification.Value, GetExpressionContext(), lExpression);
            }
            return null;
        }
    }
}
