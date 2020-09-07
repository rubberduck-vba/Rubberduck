using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.EncapsulateField;
using System.Collections.Generic;

namespace Rubberduck.Refactorings
{
    public interface IEncapsulateFieldConflictFinderFactory
    {
        IEncapsulateFieldConflictFinder CreateEncapsulateFieldUseBackingFieldConflictFinder(IReadOnlyCollection<IEncapsulateFieldCandidate> candidates);
        IEncapsulateFieldConflictFinder CreateEncapsulateFieldUseBackingUDTMemberConflictFinder(IReadOnlyCollection<IEncapsulateFieldCandidate> candidates, IReadOnlyCollection<IObjectStateUDT> objectStateUDTs);
    }

    public class EncapsulateFieldConflictFinderFactory : IEncapsulateFieldConflictFinderFactory
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        public EncapsulateFieldConflictFinderFactory(IDeclarationFinderProvider declarationFinderProvider)
        {
            _declarationFinderProvider = declarationFinderProvider;
        }

        public IEncapsulateFieldConflictFinder CreateEncapsulateFieldUseBackingFieldConflictFinder(IReadOnlyCollection<IEncapsulateFieldCandidate> candidates)
        {
            return new EncapsulateFieldUseBackingFieldsConflictFinder(_declarationFinderProvider, candidates);
        }

        public IEncapsulateFieldConflictFinder CreateEncapsulateFieldUseBackingUDTMemberConflictFinder(IReadOnlyCollection<IEncapsulateFieldCandidate> candidates, IReadOnlyCollection<IObjectStateUDT> objectStateUDTs)
        {
            return new EncapsulateFieldUseBackingUDTMemberConflictFinder(_declarationFinderProvider, candidates, objectStateUDTs);
        }
    }
}
