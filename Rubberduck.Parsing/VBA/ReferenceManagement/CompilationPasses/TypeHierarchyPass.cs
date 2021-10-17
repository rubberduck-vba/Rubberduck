using System.Collections.Generic;
using System.Linq;
using NLog;
using Rubberduck.Parsing.Binding;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA.ReferenceManagement.CompilationPasses
{
    public sealed class TypeHierarchyPass : ICompilationPass
    {
        private readonly DeclarationFinder _declarationFinder;
        private readonly BindingService _bindingService;
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
            _expressionParser = expressionParser;
        }

        public void Execute(IReadOnlyCollection<QualifiedModuleName> modules)
        {
            var toResolveSupertypesFor = _declarationFinder.UserDeclarations(DeclarationType.ClassModule)
                                            .Concat(_declarationFinder.UserDeclarations(DeclarationType.Document))
                                            .Concat(_declarationFinder.UserDeclarations(DeclarationType.UserForm))
                                            .Where(decl => modules.Contains(decl.QualifiedName.QualifiedModuleName))
                                            .Concat(_declarationFinder.BuiltInDeclarations(DeclarationType.ClassModule));
            foreach (var declaration in toResolveSupertypesFor)
            {
                AddImplementedInterface(declaration);
            }
        }

        private void AddImplementedInterface(Declaration potentialClassModule)
        {
            if (!potentialClassModule.DeclarationType.HasFlag(DeclarationType.ClassModule))
            {
                return;
            }
            var classModule = (ClassModuleDeclaration)potentialClassModule;
            foreach (var implementedInterfaceName in classModule.SupertypeNames)
            {
                if (!TrySanitizeName(implementedInterfaceName, out var sanitizedName))
                {
                    Logger.Warn("The interface name '{0}' is insane. Therefore, it cannot be resolved reliably and will be skipped.", implementedInterfaceName);
                    continue;
                }

                var expressionContext = _expressionParser.Parse(sanitizedName);
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

        private static bool TrySanitizeName(string implementedInterfaceName, out string sanitizedName)
        {
            sanitizedName = string.Empty;
            foreach (var part in implementedInterfaceName.Split('.'))
            {
                if (!string.IsNullOrWhiteSpace(sanitizedName))
                {
                    sanitizedName += ".";
                }

                if (part.StartsWith("[") && part.EndsWith("]"))
                {
                    sanitizedName += part;
                    continue;
                }

                if (part.Contains("[") || part.Contains("]"))
                {
                    sanitizedName = null;
                    break;
                }

                sanitizedName += "[" + part + "]";
            }

            return !string.IsNullOrWhiteSpace(sanitizedName);
        }
    }
}
