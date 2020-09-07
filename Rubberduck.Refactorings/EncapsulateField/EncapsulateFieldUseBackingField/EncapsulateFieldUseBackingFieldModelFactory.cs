using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
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
        /// <param name="clientTarget">Optional: <c>UserDefinedType</c> Field to include the Encapsulated Field(s)</param>
        EncapsulateFieldUseBackingFieldModel Create(IEnumerable<EncapsulateFieldRequest> requests);

        /// <summary>
        /// Creates an <c>EncapsulateFieldUseBackingFieldModel</c> based upon collection of
        /// <c>IEncapsulateFieldCandidate</c> instances created by <c>EncapsulateFieldCandidateCollectionFactory</c>.  
        /// This function is intended for exclusive use by <c>EncapsulateFieldModelFactory</c>
        /// </summary>
        EncapsulateFieldUseBackingFieldModel Create(IReadOnlyCollection<IEncapsulateFieldCandidate> candidates, IEnumerable<EncapsulateFieldRequest> requests);
    }

    public class EncapsulateFieldUseBackingFieldModelFactory : IEncapsulateFieldUseBackingFieldModelFactory
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IEncapsulateFieldCandidateCollectionFactory _fieldCandidateCollectionFactory;
        private readonly IEncapsulateFieldConflictFinderFactory _conflictFinderFactory;

        public EncapsulateFieldUseBackingFieldModelFactory(IDeclarationFinderProvider declarationFinderProvider, 
            IEncapsulateFieldCandidateCollectionFactory encapsulateFieldCandidateCollectionFactory,
            IEncapsulateFieldConflictFinderFactory encapsulateFieldConflictFinderFactory)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _fieldCandidateCollectionFactory = encapsulateFieldCandidateCollectionFactory;
            _conflictFinderFactory = encapsulateFieldConflictFinderFactory;
        }

        public EncapsulateFieldUseBackingFieldModel Create(IEnumerable<EncapsulateFieldRequest> requests)
        {
            if (!requests.Any())
            {
                return new EncapsulateFieldUseBackingFieldModel(Enumerable.Empty<IEncapsulateFieldCandidate>(), _declarationFinderProvider);
            }

            var fieldCandidates = _fieldCandidateCollectionFactory.Create(requests.First().Declaration.QualifiedModuleName);
            return Create(fieldCandidates, requests);
        }

        public EncapsulateFieldUseBackingFieldModel Create(IReadOnlyCollection<IEncapsulateFieldCandidate> candidates, IEnumerable<EncapsulateFieldRequest> requests)
        {
            var fieldCandidates = candidates.ToList();

            foreach (var request in requests)
            {
                var candidate = fieldCandidates.Single(c => c.Declaration.Equals(request.Declaration));
                request.ApplyRequest(candidate);
            }

            var conflictsFinder = _conflictFinderFactory.CreateEncapsulateFieldUseBackingFieldConflictFinder(fieldCandidates);
            fieldCandidates.ForEach(c => c.ConflictFinder = conflictsFinder);

            return new EncapsulateFieldUseBackingFieldModel(fieldCandidates, _declarationFinderProvider)
            {
                ConflictFinder = conflictsFinder
            };
        }
    }
}
