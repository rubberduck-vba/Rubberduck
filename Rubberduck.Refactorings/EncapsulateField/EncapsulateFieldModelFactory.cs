using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.EncapsulateField;
using System;
using System.Collections.Generic;

namespace Rubberduck.Refactorings
{
    public interface IEncapsulateFieldModelFactory
    {
        /// <summary>
        /// Creates the supporting EncapsulateFieldRefactoringAction models for the EncapsulateFieldRefactoring.
        /// </summary>
        EncapsulateFieldModel Create(Declaration target);
    }

    public class EncapsulateFieldModelFactory : IEncapsulateFieldModelFactory
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IEncapsulateFieldCandidateFactory _candidatesFactory;
        private readonly IEncapsulateFieldUseBackingUDTMemberModelFactory _useBackingUDTMemberModelFactory;
        private readonly IEncapsulateFieldUseBackingFieldModelFactory _useBackingFieldModelFactory;
        private readonly IEncapsulateFieldCandidateSetsProviderFactory _candidateSetsFactory;
        private readonly IEncapsulateFieldConflictFinderFactory _encapsulateFieldConflictFinderFactory;

        public EncapsulateFieldModelFactory(IDeclarationFinderProvider declarationFinderProvider,
            IEncapsulateFieldCandidateFactory candidatesFactory,
            IEncapsulateFieldUseBackingUDTMemberModelFactory encapsulateFieldUseBackingUDTMemberModelFactory,
            IEncapsulateFieldUseBackingFieldModelFactory encapsulateFieldUseBackingFieldModelFactory,
            IEncapsulateFieldCandidateSetsProviderFactory candidateSetsProviderFactory,
            IEncapsulateFieldConflictFinderFactory encapsulateFieldConflictFinderFactory)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _candidatesFactory = candidatesFactory;
            _useBackingUDTMemberModelFactory = encapsulateFieldUseBackingUDTMemberModelFactory as IEncapsulateFieldUseBackingUDTMemberModelFactory;
            _useBackingFieldModelFactory = encapsulateFieldUseBackingFieldModelFactory;
            _candidateSetsFactory = candidateSetsProviderFactory;
            _encapsulateFieldConflictFinderFactory = encapsulateFieldConflictFinderFactory;
        }

        public EncapsulateFieldModel Create(Declaration target)
        {
            if (!(target is VariableDeclaration targetField))
            {
                throw new ArgumentException();
            }

            var fieldEncapsulationModels = new List<FieldEncapsulationModel>()
            {
                new FieldEncapsulationModel(targetField)
            };

            var contextCollections = _candidateSetsFactory.Create(_declarationFinderProvider, _candidatesFactory, target.QualifiedModuleName);

            var useBackingFieldModel = _useBackingFieldModelFactory.Create(contextCollections, fieldEncapsulationModels);
            var useBackingUDTMemberModel = _useBackingUDTMemberModelFactory.Create(contextCollections, fieldEncapsulationModels);

            var initialStrategy = useBackingUDTMemberModel.ObjectStateUDTField.IsExistingDeclaration
                ? EncapsulateFieldStrategy.ConvertFieldsToUDTMembers
                : EncapsulateFieldStrategy.UseBackingFields;

            var conflictFinder = _encapsulateFieldConflictFinderFactory.Create(_declarationFinderProvider,
                contextCollections.EncapsulateFieldUseBackingFieldCandidates,
                contextCollections.ObjectStateFieldCandidates);

            var model = new EncapsulateFieldModel(useBackingFieldModel, useBackingUDTMemberModel, conflictFinder)
            {
                EncapsulateFieldStrategy = initialStrategy,
            };

            return model;
        }
    }
}
