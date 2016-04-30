using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public sealed class SimpleNameDefaultBinding : IExpressionBinding
    {
        private readonly DeclarationFinder _declarationFinder;
        private readonly Declaration _project;
        private readonly Declaration _module;
        private readonly Declaration _parent;
        private readonly VBAExpressionParser.SimpleNameExpressionContext _expression;

        public SimpleNameDefaultBinding(
            DeclarationFinder declarationFinder,
            Declaration module, 
            Declaration parent, 
            VBAExpressionParser.SimpleNameExpressionContext expression)
        {
            _declarationFinder = declarationFinder;
            _project = module.ParentDeclaration;
            _module = module;
            _parent = parent;
            _expression = expression;
        }

        public IBoundExpression Resolve()
        {
            string name = ExpressionName.GetName(_expression.name());
            IBoundExpression boundExpression = null;
            boundExpression = ResolveProcedureNamespace(name);
            if (boundExpression != null)
            {
                return boundExpression;
            }
            boundExpression = ResolveEnclosingModuleNamespace(name);
            if (boundExpression != null)
            {
                return boundExpression;
            }
            boundExpression = ResolveEnclosingProjectNamespace(name);
            if (boundExpression != null)
            {
                return boundExpression;
            }
            boundExpression = ResolveOtherProceduralModuleEnclosingProjectNamespace(name);
            if (boundExpression != null)
            {
                return boundExpression;
            }
            boundExpression = ResolveReferencedProjectNamespace(name);
            if (boundExpression != null)
            {
                return boundExpression;
            }
            boundExpression = ResolveModuleReferencedProjectNamespace(name);
            if (boundExpression != null)
            {
                return boundExpression;
            }
            return null;
        }

        private IBoundExpression ResolveProcedureNamespace(string name)
        {
            if (_parent.DeclarationType != DeclarationType.Function && _parent.DeclarationType != DeclarationType.Procedure)
            {
                return null;
            }
            /*  Namespace tier 1:
                Procedure namespace: A local variable, reference parameter binding or constant whose implicit 
                or explicit definition precedes this expression in an enclosing procedure.                
            */
            var localVariable = _declarationFinder.FindMemberEnclosingProcedure(_parent, name, DeclarationType.Variable);
            if (localVariable != null)
            {
                return new SimpleNameExpression(localVariable, ExpressionClassification.Variable, _expression);
            }
            //
            var parameter = _declarationFinder.FindMemberEnclosingProcedure(_parent, name, DeclarationType.Parameter);
            if (parameter != null)
            {
                return new SimpleNameExpression(parameter, ExpressionClassification.Variable, _expression);
            }
            var constant = _declarationFinder.FindMemberEnclosingProcedure(_parent, name, DeclarationType.Constant);
            if (constant != null)
            {
                return new SimpleNameExpression(constant, ExpressionClassification.Value, _expression);
            }
            return null;
        }

        private IBoundExpression ResolveEnclosingModuleNamespace(string name)
        {
            /*  Namespace tier 2:
                Enclosing Module namespace: A variable, constant, Enum type, Enum member, property, 
                function or subroutine defined at the module-level in the enclosing module.
            */
            var moduleVariable = _declarationFinder.FindMemberEnclosingModule(_project, _module, _parent, name, DeclarationType.Variable);
            if (moduleVariable != null)
            {
                return new SimpleNameExpression(moduleVariable, ExpressionClassification.Variable, _expression);
            }
            var moduleConstant = _declarationFinder.FindMemberEnclosingModule(_project, _module, _parent, name, DeclarationType.Constant);
            if (moduleConstant != null)
            {
                return new SimpleNameExpression(moduleConstant, ExpressionClassification.Variable, _expression);
            }
            var enumType = _declarationFinder.FindMemberEnclosingModule(_project, _module, _parent, name, DeclarationType.Enumeration);
            if (enumType != null)
            {
                return new SimpleNameExpression(enumType, ExpressionClassification.Type, _expression);
            }
            var enumMember = _declarationFinder.FindMemberEnclosingModule(_project, _module, _parent, name, DeclarationType.EnumerationMember);
            if (enumMember != null)
            {
                return new SimpleNameExpression(enumMember, ExpressionClassification.Value, _expression);
            }
            var property = _declarationFinder.FindMemberEnclosingModule(_project, _module, _parent, name, DeclarationType.Property);
            if (property != null)
            {
                return new SimpleNameExpression(property, ExpressionClassification.Property, _expression);
            }
            var function = _declarationFinder.FindMemberEnclosingModule(_project, _module, _parent, name, DeclarationType.Function);
            if (function != null)
            {
                return new SimpleNameExpression(function, ExpressionClassification.Function, _expression);
            }
            var subroutine = _declarationFinder.FindMemberEnclosingModule(_project, _module, _parent, name, DeclarationType.Procedure);
            if (subroutine != null)
            {
                return new SimpleNameExpression(subroutine, ExpressionClassification.Subroutine, _expression);
            }
            return null;
        }

        private IBoundExpression ResolveEnclosingProjectNamespace(string name)
        {
            /*  Namespace tier 3:
                Enclosing Project namespace: The enclosing project itself, a referenced project, or a 
                procedural module contained in the enclosing project.
            */
            if (_declarationFinder.IsMatch(_project.Project.Name, name))
            {
                return new SimpleNameExpression(_project, ExpressionClassification.Project, _expression);
            }
            var referencedProject = _declarationFinder.FindReferencedProject(_project, name);
            if (referencedProject != null)
            {
                return new SimpleNameExpression(referencedProject, ExpressionClassification.Project, _expression);
            }
            var proceduralModuleEnclosingProject = _declarationFinder.FindModuleEnclosingProjectWithoutEnclosingModule(_project, _module, name, DeclarationType.ProceduralModule);
            if (proceduralModuleEnclosingProject != null)
            {
                return new SimpleNameExpression(proceduralModuleEnclosingProject, ExpressionClassification.ProceduralModule, _expression);
            }
            return null;
        }

        private IBoundExpression ResolveOtherProceduralModuleEnclosingProjectNamespace(string name)
        {
            /*  Namespace tier 4:
                Other Procedural Module in Enclosing Project namespace: An accessible variable, constant, 
                Enum type, Enum member, property, function or subroutine defined in a procedural module 
                within the enclosing project other than the enclosing module.  
            */
            var accessibleVariable = _declarationFinder.FindMemberEnclosedProjectWithoutEnclosingModule(_project, _module, _parent, name, DeclarationType.Variable);
            if (accessibleVariable != null)
            {
                return new SimpleNameExpression(accessibleVariable, ExpressionClassification.Variable, _expression);
            }
            var accessibleConstant = _declarationFinder.FindMemberEnclosedProjectWithoutEnclosingModule(_project, _module, _parent, name, DeclarationType.Constant);
            if (accessibleConstant != null)
            {
                return new SimpleNameExpression(accessibleConstant, ExpressionClassification.Variable, _expression);
            }
            var accessibleType = _declarationFinder.FindMemberEnclosedProjectWithoutEnclosingModule(_project, _module, _parent, name, DeclarationType.Enumeration);
            if (accessibleType != null)
            {
                return new SimpleNameExpression(accessibleType, ExpressionClassification.Type, _expression);
            }
            var accessibleMember = _declarationFinder.FindMemberEnclosedProjectWithoutEnclosingModule(_project, _module, _parent, name, DeclarationType.EnumerationMember);
            if (accessibleMember != null)
            {
                return new SimpleNameExpression(accessibleMember, ExpressionClassification.Value, _expression);
            }
            var accessibleProperty = _declarationFinder.FindMemberEnclosedProjectWithoutEnclosingModule(_project, _module, _parent, name, DeclarationType.Property);
            if (accessibleProperty != null)
            {
                return new SimpleNameExpression(accessibleProperty, ExpressionClassification.Property, _expression);
            }
            var accessibleFunction = _declarationFinder.FindMemberEnclosedProjectWithoutEnclosingModule(_project, _module, _parent, name, DeclarationType.Function);
            if (accessibleFunction != null)
            {
                return new SimpleNameExpression(accessibleFunction, ExpressionClassification.Function, _expression);
            }
            var accessibleSubroutine = _declarationFinder.FindMemberEnclosedProjectWithoutEnclosingModule(_project, _module, _parent, name, DeclarationType.Procedure);
            if (accessibleSubroutine != null)
            {
                return new SimpleNameExpression(accessibleSubroutine, ExpressionClassification.Subroutine, _expression);
            }
            return null;
        }

        private IBoundExpression ResolveReferencedProjectNamespace(string name)
        {
            /*  Namespace tier 5:
                Referenced Project namespace: An accessible procedural module contained in a referenced 
                project.
            */
            var accessibleModule = _declarationFinder.FindModuleReferencedProject(_project, _module, name, DeclarationType.ProceduralModule);
            if (accessibleModule != null)
            {
                return new SimpleNameExpression(accessibleModule, ExpressionClassification.ProceduralModule, _expression);
            }
            return null;
        }

        private IBoundExpression ResolveModuleReferencedProjectNamespace(string name)
        {
            /*  Namespace tier 6:
                Module in Referenced Project namespace: An accessible variable, constant, Enum type, 
                Enum member, property, function or subroutine defined in a procedural module or as a member 
                of the default instance of a global class module within a referenced project.  
            */

            // Part 1: Procedural module as parent
            var accessibleVariable = _declarationFinder.FindMemberReferencedProjectInModule(_project, _module, _parent, DeclarationType.ProceduralModule, name, DeclarationType.Variable);
            if (accessibleVariable != null)
            {
                return new SimpleNameExpression(accessibleVariable, ExpressionClassification.Variable, _expression);
            }
            var accessibleConstant = _declarationFinder.FindMemberReferencedProjectInModule(_project, _module, _parent, DeclarationType.ProceduralModule, name, DeclarationType.Constant);
            if (accessibleConstant != null)
            {
                return new SimpleNameExpression(accessibleConstant, ExpressionClassification.Variable, _expression);
            }
            var accessibleType = _declarationFinder.FindMemberReferencedProjectInModule(_project, _module, _parent, DeclarationType.ProceduralModule, name, DeclarationType.Enumeration);
            if (accessibleType != null)
            {
                return new SimpleNameExpression(accessibleType, ExpressionClassification.Type, _expression);
            }
            var accessibleMember = _declarationFinder.FindMemberReferencedProjectInModule(_project, _module, _parent, DeclarationType.ProceduralModule, name, DeclarationType.EnumerationMember);
            if (accessibleMember != null)
            {
                return new SimpleNameExpression(accessibleMember, ExpressionClassification.Value, _expression);
            }
            var accessibleProperty = _declarationFinder.FindMemberReferencedProjectInModule(_project, _module, _parent, DeclarationType.ProceduralModule, name, DeclarationType.Property);
            if (accessibleProperty != null)
            {
                return new SimpleNameExpression(accessibleProperty, ExpressionClassification.Property, _expression);
            }
            var accessibleFunction = _declarationFinder.FindMemberReferencedProjectInModule(_project, _module, _parent, DeclarationType.ProceduralModule, name, DeclarationType.Function);
            if (accessibleFunction != null)
            {
                return new SimpleNameExpression(accessibleFunction, ExpressionClassification.Function, _expression);
            }
            var accessibleSubroutine = _declarationFinder.FindMemberReferencedProjectInModule(_project, _module, _parent, DeclarationType.ProceduralModule, name, DeclarationType.Procedure);
            if (accessibleSubroutine != null)
            {
                return new SimpleNameExpression(accessibleSubroutine, ExpressionClassification.Subroutine, _expression);
            }

            // Part 2: Global class module as parent
            var globalClassModuleVariable = _declarationFinder.FindMemberReferencedProjectInGlobalClassModule(_project, _module, _parent, name, DeclarationType.Variable);
            if (globalClassModuleVariable != null)
            {
                return new SimpleNameExpression(globalClassModuleVariable, ExpressionClassification.Variable, _expression);
            }
            var globalClassModuleConstant = _declarationFinder.FindMemberReferencedProjectInGlobalClassModule(_project, _module, _parent, name, DeclarationType.Constant);
            if (globalClassModuleConstant != null)
            {
                return new SimpleNameExpression(globalClassModuleConstant, ExpressionClassification.Variable, _expression);
            }
            var globalClassModuleType = _declarationFinder.FindMemberReferencedProjectInGlobalClassModule(_project, _module, _parent, name, DeclarationType.Enumeration);
            if (globalClassModuleType != null)
            {
                return new SimpleNameExpression(globalClassModuleType, ExpressionClassification.Type, _expression);
            }
            var globalClassModuleMember = _declarationFinder.FindMemberReferencedProjectInGlobalClassModule(_project, _module, _parent, name, DeclarationType.EnumerationMember);
            if (globalClassModuleMember != null)
            {
                return new SimpleNameExpression(globalClassModuleMember, ExpressionClassification.Value, _expression);
            }
            var globalClassModuleProperty = _declarationFinder.FindMemberReferencedProjectInGlobalClassModule(_project, _module, _parent, name, DeclarationType.Property);
            if (globalClassModuleProperty != null)
            {
                return new SimpleNameExpression(globalClassModuleProperty, ExpressionClassification.Property, _expression);
            }
            var globalClassModuleFunction = _declarationFinder.FindMemberReferencedProjectInGlobalClassModule(_project, _module, _parent, name, DeclarationType.Function);
            if (globalClassModuleFunction != null)
            {
                return new SimpleNameExpression(globalClassModuleFunction, ExpressionClassification.Function, _expression);
            }
            var globalClassModuleSubroutine = _declarationFinder.FindMemberReferencedProjectInGlobalClassModule(_project, _module, _parent, name, DeclarationType.Procedure);
            if (globalClassModuleSubroutine != null)
            {
                return new SimpleNameExpression(globalClassModuleSubroutine, ExpressionClassification.Subroutine, _expression);
            }
            return null;
        }
    }
}
