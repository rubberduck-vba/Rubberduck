using NLog;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Binding;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Parsing.Symbols
{
    public sealed class TypeHierarchyPass : ICompilationPass
    {
        private readonly DeclarationFinder _declarationFinder;
        private readonly BindingService _bindingService;
        private readonly BoundExpressionVisitor _boundExpressionVisitor;
        private readonly VBAExpressionParser _expressionParser;
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

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

        public void Execute(IReadOnlyCollection<QualifiedModuleName> modules)
        {
            var toRelsolveSupertypesFor = _declarationFinder
                                            .UserDeclarations(DeclarationType.ClassModule)
                                            .Where(decl => modules.Contains(decl.QualifiedName.QualifiedModuleName))
                                            .Concat(_declarationFinder.BuiltInDeclarations(DeclarationType.ClassModule));
            foreach (var declaration in toRelsolveSupertypesFor)
            {
                AddImplementedInterface(declaration);
            }
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
                if (implementedInterface.Classification != ExpressionClassification.ResolutionFailed)
                {
                    classModule.AddSupertype(implementedInterface.ReferencedDeclaration);
                }
                else
                {
                    Logger.Warn("Failed to resolve interface {0}.", implementedInterfaceName);
                }
            }
        }
    }
}
