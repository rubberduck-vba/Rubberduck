using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.Refactorings.EncapsulateFieldUseBackingField;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings
{
    public interface IEncapsulateFieldUseBackingFieldModelFactory
    {
        /// <summary>
        /// Creates an <c>EncapsulateFieldUseBackingFieldModel</c> used by the <c>EncapsulateFieldUseBackingFieldRefactoringAction</c>.
        /// </summary>
        EncapsulateFieldUseBackingFieldModel Create(IEnumerable<FieldEncapsulationModel> requests);

        /// <summary>
        /// Creates an <c>EncapsulateFieldUseBackingFieldModel</c> based upon collection of
        /// <c>IEncapsulateFieldCandidate</c> instances created by <c>EncapsulateFieldCandidateCollectionFactory</c>.  
        /// This function is intended for exclusive use by the <c>EncapsulateFieldModelFactory</c>
        /// </summary>
        EncapsulateFieldUseBackingFieldModel Create(IEncapsulateFieldCollectionsProvider collectionsProvider, IEnumerable<FieldEncapsulationModel> requests);
    }

    public class EncapsulateFieldUseBackingFieldModelFactory : IEncapsulateFieldUseBackingFieldModelFactory
    {
        private readonly IEncapsulateFieldCollectionsProviderFactory _collectionProviderFactory;
        private readonly IEncapsulateFieldConflictFinderFactory _conflictFinderFactory;

        public EncapsulateFieldUseBackingFieldModelFactory(
            IEncapsulateFieldCollectionsProviderFactory encapsulateFieldCollectionsProviderFactory,
            IEncapsulateFieldConflictFinderFactory encapsulateFieldConflictFinderFactory)
        {
            _conflictFinderFactory = encapsulateFieldConflictFinderFactory;
            _collectionProviderFactory = encapsulateFieldCollectionsProviderFactory;
        }

        public EncapsulateFieldUseBackingFieldModel Create(IEnumerable<FieldEncapsulationModel> requests)
        {
            if (!requests.Any())
            {
                return new EncapsulateFieldUseBackingFieldModel(Enumerable.Empty<IEncapsulateFieldCandidate>());
            }

            var collectionsProvider = _collectionProviderFactory.Create(requests.First().Declaration.QualifiedModuleName);
            return Create(collectionsProvider, requests);
        }

        public EncapsulateFieldUseBackingFieldModel Create(IEncapsulateFieldCollectionsProvider collectionsProvider, IEnumerable<FieldEncapsulationModel> requests)
        {
            var fieldCandidates = collectionsProvider.EncapsulateFieldCandidates.ToList();

            foreach (var request in requests)
            {
                var candidate = fieldCandidates.Single(c => c.Declaration.Equals(request.Declaration));
                candidate.EncapsulateFlag = true;
                candidate.IsReadOnly = request.IsReadOnly;
                if (request.PropertyIdentifier != null)
                {
                    candidate.PropertyIdentifier = request.PropertyIdentifier;
                }
            }

            var conflictsFinder = _conflictFinderFactory.Create(collectionsProvider);

            fieldCandidates.ForEach(c => c.ConflictFinder = conflictsFinder);

            return new EncapsulateFieldUseBackingFieldModel(fieldCandidates)
            {
                ConflictFinder = conflictsFinder
            };
        }
    }
}
