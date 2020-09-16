using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.EncapsulateField;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings
{
    public interface IEncapsulateFieldConflictFinderFactory
    {
        IEncapsulateFieldConflictFinder Create(IEncapsulateFieldCollectionsProvider collectionProvider);
    }

    public class EncapsulateFieldConflictFinderFactory : IEncapsulateFieldConflictFinderFactory
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        public EncapsulateFieldConflictFinderFactory(IDeclarationFinderProvider declarationFinderProvider)
        {
            _declarationFinderProvider = declarationFinderProvider;
        }

        public IEncapsulateFieldConflictFinder Create(IEncapsulateFieldCollectionsProvider collectionProvider)
        {
            return new EncapsulateFieldConflictFinder(_declarationFinderProvider, collectionProvider.EncapsulateFieldCandidates, collectionProvider.ObjectStateUDTCandidates);
        }
    }
}
