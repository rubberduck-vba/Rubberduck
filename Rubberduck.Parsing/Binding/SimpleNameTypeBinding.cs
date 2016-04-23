using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public sealed class SimpleNameTypeBinding : IExpressionBinding
    {
        private readonly DeclarationFinder _declarationFinder;
        private readonly Declaration _module;
        private readonly VBAExpressionParser.SimpleNameExpressionContext _expression;

        public SimpleNameTypeBinding(DeclarationFinder declarationFinder, Declaration module, VBAExpressionParser.SimpleNameExpressionContext expression)
        {
            _declarationFinder = declarationFinder;
            _module = module;
            _expression = expression;
        }

        public IBoundExpression Resolve()
        {
            IBoundExpression boundExpression = null;
            string name = ExpressionName.GetName(_expression.name());
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
            return boundExpression;
        }

        private IBoundExpression ResolveEnclosingModule(string name)
        {            
            /*  Namespace tier 1:
                Enclosing Module namespace: A UDT or Enum type defined at the module-level in the 
                enclosing module.
            */
            var udt = _declarationFinder.Find(_module, name, DeclarationType.UserDefinedType);
            if (udt != null)
            {
                return new SimpleNameExpression(udt, ExpressionClassification.Type, _expression);
            }
            var enumType = _declarationFinder.Find(_module, name, DeclarationType.Enumeration);
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
            var enclosingProjectDeclaration = _module.ParentDeclaration;
            if (enclosingProjectDeclaration.Project.Name == name)
            {
                return new SimpleNameExpression(enclosingProjectDeclaration, ExpressionClassification.Project, _expression);
            }
            var referencedProject = _declarationFinder.FindReferencedProject(enclosingProjectDeclaration, name);
            if (referencedProject != null)
            {
                return new SimpleNameExpression(referencedProject, ExpressionClassification.Type, _expression);
            }
            var proceduralModuleEnclosingProject = _declarationFinder.Find(enclosingProjectDeclaration, name, DeclarationType.ProceduralModule);
            if (proceduralModuleEnclosingProject != null)
            {
                return new SimpleNameExpression(proceduralModuleEnclosingProject, ExpressionClassification.ProceduralModule, _expression);
            }
            var classEnclosingProject = _declarationFinder.Find(enclosingProjectDeclaration, name, DeclarationType.ClassModule);
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
            Declaration enclosingProjectDeclaration = _module.ParentDeclaration;
            var accessibleUdt = _declarationFinder.FindAccessibleInEnclosingProject(enclosingProjectDeclaration, _module, name, DeclarationType.UserDefinedType);
            if (accessibleUdt != null)
            {
                return new SimpleNameExpression(accessibleUdt, ExpressionClassification.Type, _expression);
            }
            var accessibleType = _declarationFinder.FindAccessibleInEnclosingProject(enclosingProjectDeclaration, _module, name, DeclarationType.Enumeration);
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
            var enclosingProjectDeclaration = _module.ParentDeclaration;
            var accessibleModule = _declarationFinder.FindProceduralModuleInReferencedProject(enclosingProjectDeclaration, name);
            if (accessibleModule != null)
            {
                return new SimpleNameExpression(accessibleModule, ExpressionClassification.ProceduralModule, _expression);
            }
            var accessibleClass = _declarationFinder.FindClassModuleInReferencedProject(enclosingProjectDeclaration, name);
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
            var enclosingProjectDeclaration = _module.ParentDeclaration;
            var referencedProjectUdt = _declarationFinder.FindTypeInReferencedProject(enclosingProjectDeclaration, name, DeclarationType.UserDefinedType);
            if (referencedProjectUdt != null)
            {
                return new SimpleNameExpression(referencedProjectUdt, ExpressionClassification.Type, _expression);
            }
            var referencedProjectEnumType = _declarationFinder.FindTypeInReferencedProject(enclosingProjectDeclaration, name, DeclarationType.Enumeration);
            if (referencedProjectEnumType != null)
            {
                return new SimpleNameExpression(referencedProjectEnumType, ExpressionClassification.Type, _expression);
            }
            return null;
        }
    }
}
