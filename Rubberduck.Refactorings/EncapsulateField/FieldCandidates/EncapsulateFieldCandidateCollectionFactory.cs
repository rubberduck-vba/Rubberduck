using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings
{
    public interface IEncapsulateFieldCandidateCollectionFactory
    {
        IReadOnlyCollection<IEncapsulateFieldCandidate> Create(QualifiedModuleName qualifiedModuleName);
    }

    public class EncapsulateFieldCandidateCollectionFactory : IEncapsulateFieldCandidateCollectionFactory
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IEncapsulateFieldCandidateFactory _fieldCandidateFactory;
        public EncapsulateFieldCandidateCollectionFactory(
            IDeclarationFinderProvider declarationFinderProvider,
            IEncapsulateFieldCandidateFactory encapsulateFieldCandidateFactory)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _fieldCandidateFactory = encapsulateFieldCandidateFactory;
        }

        public IReadOnlyCollection<IEncapsulateFieldCandidate> Create(QualifiedModuleName qualifiedModuleName)
        {
            return _declarationFinderProvider.DeclarationFinder.Members(qualifiedModuleName, DeclarationType.Variable)
                .Where(v => v.ParentDeclaration is ModuleDeclaration
                    && !v.IsWithEvents)
                .Select(f => _fieldCandidateFactory.Create(f))
                .ToList();
        }
    }
}
