using Rubberduck.Parsing.Symbols;
using System.Linq;

namespace Rubberduck.Parsing.Binding
{
    public sealed class SimpleNameDefaultBinding : IExpressionBinding
    {
        private readonly DeclarationFinder _declarationFinder;
        private readonly Declaration _project;
        private readonly Declaration _module;
        private readonly Declaration _parent;
        private readonly VBAExpressionParser.SimpleNameExpressionContext _expression;
        private readonly DeclarationType _propertySearchType;

        public SimpleNameDefaultBinding(
            DeclarationFinder declarationFinder,
            Declaration project,
            Declaration module,
            Declaration parent,
            VBAExpressionParser.SimpleNameExpressionContext expression,
            ResolutionStatementContext statementContext)
        {
            _declarationFinder = declarationFinder;
            _project = project;
            _module = module;
            _parent = parent;
            _expression = expression;
            _propertySearchType = StatementContext.GetSearchDeclarationType(statementContext);
        }

        public bool IsPotentialLeftMatch { get; internal set; }

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
            /*  Namespace tier 1:
                Procedure namespace: A local variable, reference parameter binding or constant whose implicit 
                or explicit definition precedes this expression in an enclosing procedure.                
            */
            if (!_parent.DeclarationType.HasFlag(DeclarationType.Function) && !_parent.DeclarationType.HasFlag(DeclarationType.Procedure))
            {
                return null;
            }
            var localVariable = _declarationFinder.FindMemberEnclosingProcedure(_parent, name, DeclarationType.Variable);
            if (IsValidMatch(localVariable, name))
            {
                return new SimpleNameExpression(localVariable, ExpressionClassification.Variable, _expression);
            }
            var parameter = _declarationFinder.FindMemberEnclosingProcedure(_parent, name, DeclarationType.Parameter);
            if (IsValidMatch(parameter, name))
            {
                return new SimpleNameExpression(parameter, ExpressionClassification.Variable, _expression);
            }
            var constant = _declarationFinder.FindMemberEnclosingProcedure(_parent, name, DeclarationType.Constant);
            if (IsValidMatch(constant, name))
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
            if (IsValidMatch(moduleVariable, name))
            {
                return new SimpleNameExpression(moduleVariable, ExpressionClassification.Variable, _expression);
            }
            var moduleConstant = _declarationFinder.FindMemberEnclosingModule(_project, _module, _parent, name, DeclarationType.Constant);
            if (IsValidMatch(moduleConstant, name))
            {
                return new SimpleNameExpression(moduleConstant, ExpressionClassification.Variable, _expression);
            }
            var enumType = _declarationFinder.FindMemberEnclosingModule(_project, _module, _parent, name, DeclarationType.Enumeration);
            if (IsValidMatch(enumType, name))
            {
                return new SimpleNameExpression(enumType, ExpressionClassification.Type, _expression);
            }
            var enumMember = _declarationFinder.FindMemberEnclosingModule(_project, _module, _parent, name, DeclarationType.EnumerationMember);
            if (IsValidMatch(enumMember, name))
            {
                return new SimpleNameExpression(enumMember, ExpressionClassification.Value, _expression);
            }
            var property = _declarationFinder.FindMemberEnclosingModule(_project, _module, _parent, name, _propertySearchType);
            if (IsValidMatch(property, name))
            {
                return new SimpleNameExpression(property, ExpressionClassification.Property, _expression);
            }
            var function = _declarationFinder.FindMemberEnclosingModule(_project, _module, _parent, name, DeclarationType.Function);
            if (IsValidMatch(function, name))
            {
                return new SimpleNameExpression(function, ExpressionClassification.Function, _expression);
            }
            var subroutine = _declarationFinder.FindMemberEnclosingModule(_project, _module, _parent, name, DeclarationType.Procedure);
            if (IsValidMatch(subroutine, name))
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
            if (_declarationFinder.IsMatch(_project.ProjectName, name))
            {
                return new SimpleNameExpression(_project, ExpressionClassification.Project, _expression);
            }
            var referencedProject = _declarationFinder.FindReferencedProject(_project, name);
            if (referencedProject != null)
            {
                return new SimpleNameExpression(referencedProject, ExpressionClassification.Project, _expression);
            }
            if (_module.DeclarationType == DeclarationType.ProceduralModule && _declarationFinder.IsMatch(_module.IdentifierName, name))
            {
                return new SimpleNameExpression(_module, ExpressionClassification.ProceduralModule, _expression);
            }
            var proceduralModuleEnclosingProject = _declarationFinder.FindModuleEnclosingProjectWithoutEnclosingModule(_project, _module, name, DeclarationType.ProceduralModule);
            if (proceduralModuleEnclosingProject != null)
            {
                return new SimpleNameExpression(proceduralModuleEnclosingProject, ExpressionClassification.ProceduralModule, _expression);
            }
            var defaultInstanceVariableClass = _declarationFinder.FindDefaultInstanceVariableClassEnclosingProject(_project, _module, name);
            if (defaultInstanceVariableClass != null)
            {
                return new SimpleNameExpression(defaultInstanceVariableClass, ExpressionClassification.Type, _expression);
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
            if (IsValidMatch(accessibleVariable, name))
            {
                return new SimpleNameExpression(accessibleVariable, ExpressionClassification.Variable, _expression);
            }
            var accessibleConstant = _declarationFinder.FindMemberEnclosedProjectWithoutEnclosingModule(_project, _module, _parent, name, DeclarationType.Constant);
            if (IsValidMatch(accessibleConstant, name))
            {
                return new SimpleNameExpression(accessibleConstant, ExpressionClassification.Variable, _expression);
            }
            var accessibleType = _declarationFinder.FindMemberEnclosedProjectWithoutEnclosingModule(_project, _module, _parent, name, DeclarationType.Enumeration);
            if (IsValidMatch(accessibleType, name))
            {
                return new SimpleNameExpression(accessibleType, ExpressionClassification.Type, _expression);
            }
            var accessibleMember = _declarationFinder.FindMemberEnclosedProjectWithoutEnclosingModule(_project, _module, _parent, name, DeclarationType.EnumerationMember);
            if (IsValidMatch(accessibleMember, name))
            {
                return new SimpleNameExpression(accessibleMember, ExpressionClassification.Value, _expression);
            }
            var accessibleProperty = _declarationFinder.FindMemberEnclosedProjectWithoutEnclosingModule(_project, _module, _parent, name, _propertySearchType);
            if (IsValidMatch(accessibleProperty, name))
            {
                return new SimpleNameExpression(accessibleProperty, ExpressionClassification.Property, _expression);
            }
            var accessibleFunction = _declarationFinder.FindMemberEnclosedProjectWithoutEnclosingModule(_project, _module, _parent, name, DeclarationType.Function);
            if (IsValidMatch(accessibleFunction, name))
            {
                return new SimpleNameExpression(accessibleFunction, ExpressionClassification.Function, _expression);
            }
            var accessibleSubroutine = _declarationFinder.FindMemberEnclosedProjectWithoutEnclosingModule(_project, _module, _parent, name, DeclarationType.Procedure);
            if (IsValidMatch(accessibleSubroutine, name))
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
            var defaultInstanceVariableClass = _declarationFinder.FindDefaultInstanceVariableClassReferencedProject(_project, _module, name);
            if (defaultInstanceVariableClass != null)
            {
                return new SimpleNameExpression(defaultInstanceVariableClass, ExpressionClassification.Type, _expression);
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
            if (IsValidMatch(accessibleVariable, name))
            {
                return new SimpleNameExpression(accessibleVariable, ExpressionClassification.Variable, _expression);
            }
            var accessibleConstant = _declarationFinder.FindMemberReferencedProjectInModule(_project, _module, _parent, DeclarationType.ProceduralModule, name, DeclarationType.Constant);
            if (IsValidMatch(accessibleConstant, name))
            {
                return new SimpleNameExpression(accessibleConstant, ExpressionClassification.Variable, _expression);
            }
            var accessibleType = _declarationFinder.FindMemberReferencedProjectInModule(_project, _module, _parent, DeclarationType.ProceduralModule, name, DeclarationType.Enumeration);
            if (IsValidMatch(accessibleType, name))
            {
                return new SimpleNameExpression(accessibleType, ExpressionClassification.Type, _expression);
            }
            var accessibleMember = _declarationFinder.FindMemberReferencedProjectInModule(_project, _module, _parent, DeclarationType.ProceduralModule, name, DeclarationType.EnumerationMember);
            if (IsValidMatch(accessibleMember, name))
            {
                return new SimpleNameExpression(accessibleMember, ExpressionClassification.Value, _expression);
            }
            var accessibleProperty = _declarationFinder.FindMemberReferencedProjectInModule(_project, _module, _parent, DeclarationType.ProceduralModule, name, _propertySearchType);
            if (IsValidMatch(accessibleProperty, name))
            {
                return new SimpleNameExpression(accessibleProperty, ExpressionClassification.Property, _expression);
            }
            var accessibleFunction = _declarationFinder.FindMemberReferencedProjectInModule(_project, _module, _parent, DeclarationType.ProceduralModule, name, DeclarationType.Function);
            if (IsValidMatch(accessibleFunction, name))
            {
                return new SimpleNameExpression(accessibleFunction, ExpressionClassification.Function, _expression);
            }
            var accessibleSubroutine = _declarationFinder.FindMemberReferencedProjectInModule(_project, _module, _parent, DeclarationType.ProceduralModule, name, DeclarationType.Procedure);
            if (IsValidMatch(accessibleSubroutine, name))
            {
                return new SimpleNameExpression(accessibleSubroutine, ExpressionClassification.Subroutine, _expression);
            }

            // Part 2: Global class module as parent
            var globalClassModuleVariable = _declarationFinder.FindMemberReferencedProjectInGlobalClassModule(_project, _module, _parent, name, DeclarationType.Variable);
            if (IsValidMatch(globalClassModuleVariable, name))
            {
                return new SimpleNameExpression(globalClassModuleVariable, ExpressionClassification.Variable, _expression);
            }
            var globalClassModuleConstant = _declarationFinder.FindMemberReferencedProjectInGlobalClassModule(_project, _module, _parent, name, DeclarationType.Constant);
            if (IsValidMatch(globalClassModuleConstant, name))
            {
                return new SimpleNameExpression(globalClassModuleConstant, ExpressionClassification.Variable, _expression);
            }
            var globalClassModuleType = _declarationFinder.FindMemberReferencedProjectInGlobalClassModule(_project, _module, _parent, name, DeclarationType.Enumeration);
            if (IsValidMatch(globalClassModuleType, name))
            {
                return new SimpleNameExpression(globalClassModuleType, ExpressionClassification.Type, _expression);
            }
            var globalClassModuleMember = _declarationFinder.FindMemberReferencedProjectInGlobalClassModule(_project, _module, _parent, name, DeclarationType.EnumerationMember);
            if (IsValidMatch(globalClassModuleMember, name))
            {
                return new SimpleNameExpression(globalClassModuleMember, ExpressionClassification.Value, _expression);
            }
            var globalClassModuleProperty = _declarationFinder.FindMemberReferencedProjectInGlobalClassModule(_project, _module, _parent, name, _propertySearchType);
            if (IsValidMatch(globalClassModuleProperty, name))
            {
                return new SimpleNameExpression(globalClassModuleProperty, ExpressionClassification.Property, _expression);
            }
            var globalClassModuleFunction = _declarationFinder.FindMemberReferencedProjectInGlobalClassModule(_project, _module, _parent, name, DeclarationType.Function);
            if (IsValidMatch(globalClassModuleFunction, name))
            {
                return new SimpleNameExpression(globalClassModuleFunction, ExpressionClassification.Function, _expression);
            }
            var globalClassModuleSubroutine = _declarationFinder.FindMemberReferencedProjectInGlobalClassModule(_project, _module, _parent, name, DeclarationType.Procedure);
            if (IsValidMatch(globalClassModuleSubroutine, name))
            {
                return new SimpleNameExpression(globalClassModuleSubroutine, ExpressionClassification.Subroutine, _expression);
            }
            return null;
        }

        private bool IsValidMatch(Declaration match, string name)
        {
            if (match == null)
            {
                return false;
            }
            if (!IsPotentialLeftMatch || name.ToUpperInvariant() != "LEFT")
            {
                return true;
            }
            var functionSubroutinePropertyGet = match.DeclarationType == DeclarationType.Function
                || match.DeclarationType == DeclarationType.Procedure
                || match.DeclarationType == DeclarationType.PropertyGet;
            if (!functionSubroutinePropertyGet)
            {
                return true;
            }
            if (((IDeclarationWithParameter)match).Parameters.Count() > 0)
            {
                return true;
            }
            if (match.AsTypeName != null
                && match.AsTypeName.ToUpperInvariant() != "VARIANT"
                && match.AsTypeName.ToUpperInvariant() != "OBJECT"
                && match.AsTypeIsBaseType)
            {
                return false;
            }
            return true;
        }
    }
}
