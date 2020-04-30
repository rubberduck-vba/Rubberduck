using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.VBEditor;
using System.Linq;

namespace Rubberduck.Refactorings
{
    public interface IConflictDetectionDeclarationProxyFactory
    {
        IConflictDetectionDeclarationProxy Create(Declaration prototype);
        IConflictDetectionDeclarationProxy Create(Declaration target, string destinationModuleName, Accessibility? accessibility = null);
        IConflictDetectionDeclarationProxy Create(string identifier, DeclarationType declarationType, Accessibility accessibility, ModuleDeclaration targetModule, Declaration parentDeclaration);
    }

    public class ConflictDetectionDeclarationProxyFactory : IConflictDetectionDeclarationProxyFactory
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        public ConflictDetectionDeclarationProxyFactory(IDeclarationFinderProvider declarationFinderProvider)
        {
            _declarationFinderProvider = declarationFinderProvider;
        }

        public IConflictDetectionDeclarationProxy Create(string identifier, DeclarationType declarationType, Accessibility accessibility, ModuleDeclaration targetModule, Declaration parentDeclaration)
        {
            return new ConflictDetectionDeclarationProxy(identifier, declarationType, accessibility, targetModule, parentDeclaration);
        }

        public IConflictDetectionDeclarationProxy Create(Declaration prototype)
        {
            var targetModule = _declarationFinderProvider.DeclarationFinder.ModuleDeclaration(prototype.QualifiedModuleName) as ModuleDeclaration;
            return new ConflictDetectionDeclarationProxy(prototype, targetModule);
        }

        public IConflictDetectionDeclarationProxy Create(Declaration target, string destinationModuleName, Accessibility? accessibility = null)
        {
            var targetModule = _declarationFinderProvider.DeclarationFinder.DeclarationsWithType(DeclarationType.Module)
                                        .Where(m => m.IdentifierName.Equals(destinationModuleName, System.StringComparison.InvariantCultureIgnoreCase));

            return new ConflictDetectionDeclarationProxy(target, targetModule.SingleOrDefault() as ModuleDeclaration);
        }
    }
}
