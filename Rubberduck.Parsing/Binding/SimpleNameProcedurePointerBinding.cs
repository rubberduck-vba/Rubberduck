using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public sealed class SimpleNameProcedurePointerBinding : IExpressionBinding
    {
        private readonly DeclarationFinder _declarationFinder;
        private readonly Declaration _project;
        private readonly Declaration _module;
        private readonly Declaration _parent;
        private readonly VBAParser.SimpleNameExprContext _expression;

        public SimpleNameProcedurePointerBinding(
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

        public IBoundExpression Resolve()
        {
            IBoundExpression boundExpression = null;
            string name = Identifier.GetName(_expression.identifier());
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
            return new ResolutionFailedExpression();
        }

        private IBoundExpression ResolveEnclosingModule(string name)
        {
            /*
                Enclosing Module namespace: A function, subroutine or property with a Property Get defined 
                at the module-level in the enclosing module.
            */
            var function = _declarationFinder.FindMemberEnclosingModule(_module, _parent, name, DeclarationType.Function);
            if (function != null)
            {
                return new SimpleNameExpression(function, ExpressionClassification.Function, _expression);
            }
            var subroutine = _declarationFinder.FindMemberEnclosingModule(_module, _parent, name, DeclarationType.Procedure);
            if (subroutine != null)
            {
                return new SimpleNameExpression(subroutine, ExpressionClassification.Subroutine, _expression);
            }
            var propertyGet = _declarationFinder.FindMemberEnclosingModule(_module, _parent, name, DeclarationType.PropertyGet);
            if (propertyGet != null)
            {
                return new SimpleNameExpression(propertyGet, ExpressionClassification.Property, _expression);
            }
            return null;
        }

        private IBoundExpression ResolveEnclosingProject(string name)
        {
            /*
                Enclosing Project namespace: The enclosing project itself or a procedural module contained in 
                the enclosing project.  
            */
            if (_declarationFinder.IsMatch(_project.ProjectName, name))
            {
                return new SimpleNameExpression(_project, ExpressionClassification.Project, _expression);
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
            return null;
        }

        private IBoundExpression ResolveOtherModuleInEnclosingProject(string name)
        {
            /*
                Other Procedural Module in Enclosing Project namespace: An accessible function, 
                subroutine or property with a Property Get defined in a procedural module within the enclosing 
                project other than the enclosing module.  
            */
            var function = _declarationFinder.FindMemberEnclosedProjectWithoutEnclosingModule(_project, _module, _parent, name, DeclarationType.Function);
            if (function != null)
            {
                return new SimpleNameExpression(function, ExpressionClassification.Function, _expression);
            }
            var subroutine = _declarationFinder.FindMemberEnclosedProjectWithoutEnclosingModule(_project, _module, _parent, name, DeclarationType.Procedure);
            if (subroutine != null)
            {
                return new SimpleNameExpression(subroutine, ExpressionClassification.Subroutine, _expression);
            }
            var propertyGet = _declarationFinder.FindMemberEnclosedProjectWithoutEnclosingModule(_project, _module, _parent, name, DeclarationType.PropertyGet);
            if (propertyGet != null)
            {
                return new SimpleNameExpression(propertyGet, ExpressionClassification.Property, _expression);
            }
            return null;
        }
    }
}
