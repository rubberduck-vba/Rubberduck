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
        EncapsulateFieldUseBackingFieldModel Create(IEnumerable<FieldEncapsulationModel> fieldModels);

        /// <summary>
        /// Creates an <c>EncapsulateFieldUseBackingFieldModel</c> based upon collection of
        /// <c>IEncapsulateFieldCandidate</c> instances</c>.  
        /// This function is intended for exclusive use by the <c>EncapsulateFieldModelFactory</c>
        /// </summary>
        EncapsulateFieldUseBackingFieldModel Create(IEncapsulateFieldCandidateSetsProvider contextCollections, IEnumerable<FieldEncapsulationModel> fieldModels);
    }

    public class EncapsulateFieldUseBackingFieldModelFactory : IEncapsulateFieldUseBackingFieldModelFactory
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IEncapsulateFieldCandidateFactory _candidatesFactory;
        private readonly IEncapsulateFieldCandidateSetsProviderFactory _candidateSetsFactory;
        private readonly IEncapsulateFieldConflictFinderFactory _conflictFinderFactory;

        public EncapsulateFieldUseBackingFieldModelFactory(IDeclarationFinderProvider declarationFinderProvider,
            IEncapsulateFieldCandidateFactory candidatesFactory,
            IEncapsulateFieldCandidateSetsProviderFactory candidateSetsFactory,
            IEncapsulateFieldConflictFinderFactory encapsulateFieldConflictFinderFactory)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _candidatesFactory = candidatesFactory;
            _candidateSetsFactory = candidateSetsFactory;
            _conflictFinderFactory = encapsulateFieldConflictFinderFactory;
        }

        public EncapsulateFieldUseBackingFieldModel Create(IEnumerable<FieldEncapsulationModel> fieldModels)
        {
            if (!fieldModels.Any())
            {
                return new EncapsulateFieldUseBackingFieldModel(Enumerable.Empty<IEncapsulateFieldCandidate>());
            }

            var contextCollections = _candidateSetsFactory.Create(_declarationFinderProvider, _candidatesFactory, fieldModels.First().Declaration.QualifiedModuleName);

            return Create(contextCollections, fieldModels);
        }

        public EncapsulateFieldUseBackingFieldModel Create(IEncapsulateFieldCandidateSetsProvider contextCollections, IEnumerable<FieldEncapsulationModel> fieldModels)
        {
            var fieldCandidates = contextCollections.EncapsulateFieldUseBackingFieldCandidates.ToList();

            foreach (var fieldModel in fieldModels)
            {
                var candidate = fieldCandidates.Single(c => c.Declaration.Equals(fieldModel.Declaration));
                candidate.EncapsulateFlag = true;
                candidate.IsReadOnly = fieldModel.IsReadOnly;
                if (fieldModel.PropertyIdentifier != null)
                {
                    candidate.PropertyIdentifier = fieldModel.PropertyIdentifier;
                }
            }

            var conflictsFinder = _conflictFinderFactory.Create(_declarationFinderProvider,
                contextCollections.EncapsulateFieldUseBackingFieldCandidates,
                contextCollections.ObjectStateFieldCandidates);

            fieldCandidates.ForEach(c => c.ConflictFinder = conflictsFinder);

            return new EncapsulateFieldUseBackingFieldModel(fieldCandidates)
            {
                ConflictFinder = conflictsFinder
            };
        }
    }
}
