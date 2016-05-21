using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Binding;
using Rubberduck.Parsing.VBA;
using System.Diagnostics;

namespace Rubberduck.Parsing.Symbols
{
    public sealed class TypeHierarchyPass : ICompilationPass
    {
        private readonly DeclarationFinder _declarationFinder;
        private readonly BindingService _bindingService;
        private readonly BoundExpressionVisitor _boundExpressionVisitor;
        private readonly VBAExpressionParser _expressionParser;

        public TypeHierarchyPass(DeclarationFinder declarationFinder, VBAExpressionParser expressionParser)
        {
            _declarationFinder = declarationFinder;
            var typeBindingContext = new TypeBindingContext(_declarationFinder);
            var procedurePointerBindingContext = new ProcedurePointerBindingContext(_declarationFinder);
            _bindingService = new BindingService(
                _declarationFinder,
                new DefaultBindingContext(_declarationFinder, typeBindingContext, procedurePointerBindingContext),
                typeBindingContext,
                procedurePointerBindingContext);
            _boundExpressionVisitor = new BoundExpressionVisitor(new AnnotationService(_declarationFinder));
            _expressionParser = expressionParser;
        }

        public void Execute()
        {
            Stopwatch stopwatch = Stopwatch.StartNew();
            foreach (var declaration in _declarationFinder.FindClasses())
            {
                AddImplementedInterface(declaration);
            }
            stopwatch.Stop();
            Debug.WriteLine("Type hierarchies built in {0}ms.", stopwatch.ElapsedMilliseconds);
        }

        private void AddImplementedInterface(Declaration potentialClassModule)
        {
            if (potentialClassModule.DeclarationType != DeclarationType.ClassModule)
            {
                return;
            }
            var classModule = (ClassModuleDeclaration)potentialClassModule;
            foreach (var implementedInterfaceName in classModule.SupertypeNames)
            {
                var expressionContext = _expressionParser.Parse(implementedInterfaceName);
                var implementedInterface = _bindingService.ResolveType(potentialClassModule, potentialClassModule, expressionContext);
                if (implementedInterface != null)
                {
                    classModule.AddSupertype(implementedInterface.ReferencedDeclaration);
                    ((ClassModuleDeclaration)implementedInterface.ReferencedDeclaration).AddSubtype(classModule);
                }
                else
                {
                    Debug.WriteLine(string.Format("{0}: Failed to resolve interface {1}.", GetType().Name, implementedInterfaceName));
                }
            }
        }
    }
}
