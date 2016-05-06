using Rubberduck.Parsing.Binding;
using System;
using System.Diagnostics;

namespace Rubberduck.Parsing.Symbols
{
    public sealed class TypeHierarchyPass
    {
        private readonly DeclarationFinder _declarationFinder;
        private readonly BindingService _bindingService;
        private readonly BoundExpressionVisitor _boundExpressionVisitor;

        public TypeHierarchyPass(DeclarationFinder declarationFinder)
        {
            _declarationFinder = declarationFinder;
            var typeBindingContext = new TypeBindingContext(_declarationFinder);
            var procedurePointerBindingContext = new ProcedurePointerBindingContext(_declarationFinder);
            _bindingService = new BindingService(
                _declarationFinder,
                new DefaultBindingContext(_declarationFinder, typeBindingContext, procedurePointerBindingContext),
                typeBindingContext,
                procedurePointerBindingContext);
            _boundExpressionVisitor = new BoundExpressionVisitor();
        }

        public void Annotate()
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
            // TODO: Depending on how the responsibility of looking up built-in interfaces is split up, we might not have to do this here.
            if (potentialClassModule.IsBuiltIn)
            {
                return;
            }
            var classModule = (ClassModuleDeclaration)potentialClassModule;
            foreach (var implementedInterfaceName in classModule.SupertypeNames)
            {
                var implementedInterface = _bindingService.ResolveType(potentialClassModule, potentialClassModule, implementedInterfaceName);
                if (implementedInterface != null)
                {
                    classModule.AddSupertype(implementedInterface.ReferencedDeclaration);
                    ((ClassModuleDeclaration)implementedInterface.ReferencedDeclaration).AddSubtype(classModule);
                }
            }
        }
    }
}