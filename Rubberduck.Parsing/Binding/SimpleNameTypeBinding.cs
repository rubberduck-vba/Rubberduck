using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public sealed class SimpleNameTypeBinding : IExpressionBinding
    {
        private readonly DeclarationFinder _declarationFinder;
        private readonly Declaration _project;
        private readonly Declaration _module;
        private readonly Declaration _parent;
        private readonly VBAParser.SimpleNameExprContext _expression;

        public SimpleNameTypeBinding(
            DeclarationFinder declarationFinder,
            Declaration project,
            Declaration module,
            Declaration parent,
            VBAParser.SimpleNameExprContext expression)
        {
            _declarationFinder = declarationFinder;
            _project = project;
            _module = module;
            _parent = parent;
            _expression = expression;
        }

        public bool PreferProjectOverUdt { get; set; }

        public IBoundExpression Resolve()
        {
            var name = Identifier.GetName(_expression.identifier());
            if (PreferProjectOverUdt)
            {
                return ResolvePreferProject(name);
            }
            return ResolvePreferUdt(name);
        }

        private IBoundExpression ResolvePreferUdt(string name)
        {
            IBoundExpression boundExpression = null;
            boundExpression = ResolveEnclosingModule(name);
            if (boundExpression != null)
            {
                return boundExpression;
            }
            boundExpression = ResolveEnclosingProject(name);
            if (boundExpression != null)
            {
                return boundExpression;
            }
            boundExpression = ResolveOtherModuleInEnclosingProject(name);
            if (boundExpression != null)
            {
                return boundExpression;
            }
            boundExpression = ResolveReferencedProject(name);
            if (boundExpression != null)
            {
                return boundExpression;
            }
            boundExpression = ResolveModuleInReferencedProject(name);
            if (boundExpression != null)
            {
                return boundExpression;
            }
            return new ResolutionFailedExpression();
        }

        private IBoundExpression ResolvePreferProject(string name)
        {
            IBoundExpression boundExpression = null;
            // EnclosingProject and EnclosingModule have been switched.
            boundExpression = ResolveEnclosingProject(name);
            if (boundExpression != null)
            {
                return boundExpression;
            }
            boundExpression = ResolveEnclosingModule(name);
            if (boundExpression != null)
            {
                return boundExpression;
            }
            boundExpression = ResolveOtherModuleInEnclosingProject(name);
            if (boundExpression != null)
            {
                return boundExpression;
            }
            boundExpression = ResolveReferencedProject(name);
            if (boundExpression != null)
            {
                return boundExpression;
            }
            boundExpression = ResolveModuleInReferencedProject(name);
            return boundExpression;
        }

        private IBoundExpression ResolveEnclosingModule(string name)
        {
            /*  Namespace tier 1:
                Enclosing Module namespace: A UDT or Enum type defined at the module-level in the 
                enclosing module.
            */
            var udt = _declarationFinder.FindMemberEnclosingModule(_module, _parent, name, DeclarationType.UserDefinedType);
            if (udt != null)
            {
                return new SimpleNameExpression(udt, ExpressionClassification.Type, _expression);
            }
            var enumType = _declarationFinder.FindMemberEnclosingModule(_module, _parent, name, DeclarationType.Enumeration);
            if (enumType != null)
            {
                return new SimpleNameExpression(enumType, ExpressionClassification.Type, _expression);
            }
            return null;
        }

        private IBoundExpression ResolveEnclosingProject(string name)
        {
            /*  Namespace tier 2:
                Enclosing Project namespace: The enclosing project itself, a referenced project, or a 
                procedural module or class module contained in the enclosing project.  
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
            if (_module.DeclarationType == DeclarationType.ClassModule && _declarationFinder.IsMatch(_module.IdentifierName, name))
            {
                return new SimpleNameExpression(_module, ExpressionClassification.Type, _expression);
            }
            var classEnclosingProject = _declarationFinder.FindModuleEnclosingProjectWithoutEnclosingModule(_project, _module, name, DeclarationType.ClassModule);
            if (classEnclosingProject != null)
            {
                return new SimpleNameExpression(classEnclosingProject, ExpressionClassification.Type, _expression);
            }
            return null;
        }

        private IBoundExpression ResolveOtherModuleInEnclosingProject(string name)
        {
            /*  Namespace tier 3:
                Other Module in Enclosing Project namespace: An accessible UDT or Enum type defined in a 
                procedural module or class module within the enclosing project other than the enclosing module.  
            */
            var accessibleUdt = _declarationFinder.FindMemberEnclosedProjectWithoutEnclosingModule(_project, _module, _parent, name, DeclarationType.UserDefinedType);
            if (accessibleUdt != null)
            {
                return new SimpleNameExpression(accessibleUdt, ExpressionClassification.Type, _expression);
            }
            var accessibleType = _declarationFinder.FindMemberEnclosedProjectWithoutEnclosingModule(_project, _module, _parent, name, DeclarationType.Enumeration);
            if (accessibleType != null)
            {
                return new SimpleNameExpression(accessibleType, ExpressionClassification.Type, _expression);
            }
            return null;
        }

        private IBoundExpression ResolveReferencedProject(string name)
        {
            /*  Namespace tier 4:
                Referenced Project namespace: An accessible procedural module or class module contained in 
                a referenced project.
            */
            var accessibleModule = _declarationFinder.FindModuleReferencedProject(_project, _module, name, DeclarationType.ProceduralModule);
            if (accessibleModule != null)
            {
                return new SimpleNameExpression(accessibleModule, ExpressionClassification.ProceduralModule, _expression);
            }
            var accessibleClass = _declarationFinder.FindModuleReferencedProject(_project, _module, name, DeclarationType.ClassModule);
            if (accessibleClass != null)
            {
                return new SimpleNameExpression(accessibleClass, ExpressionClassification.Type, _expression);
            }
            return null;
        }

        private IBoundExpression ResolveModuleInReferencedProject(string name)
        {
            /*  Namespace tier 5:
                Module in Referenced Project namespace: An accessible UDT or Enum type defined in a 
                procedural module or class module within a referenced project.  
            */
            var referencedProjectUdt = _declarationFinder.FindMemberReferencedProject(_project, _module, _parent, name, DeclarationType.UserDefinedType);
            if (referencedProjectUdt != null)
            {
                return new SimpleNameExpression(referencedProjectUdt, ExpressionClassification.Type, _expression);
            }
            var referencedProjectEnumType = _declarationFinder.FindMemberReferencedProject(_project, _module, _parent, name, DeclarationType.Enumeration);
            if (referencedProjectEnumType != null)
            {
                return new SimpleNameExpression(referencedProjectEnumType, ExpressionClassification.Type, _expression);
            }
            var referencedProjectAliasType = _declarationFinder.FindMemberReferencedProject(_project, _module, _parent, name, DeclarationType.ComAlias);
            if (referencedProjectAliasType != null)
            {
                return new SimpleNameExpression(referencedProjectAliasType, ExpressionClassification.Type, _expression);
            }
            return null;
        }
    }
}
