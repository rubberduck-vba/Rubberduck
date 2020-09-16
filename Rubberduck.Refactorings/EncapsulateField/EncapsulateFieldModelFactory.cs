using Rubberduck.Parsing.Symbols;
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
        private readonly IEncapsulateFieldUseBackingUDTMemberModelFactory _useBackingUDTMemberModelFactory;
        private readonly IEncapsulateFieldUseBackingFieldModelFactory _useBackingFieldModelFactory;
        private readonly IEncapsulateFieldCollectionsProviderFactory _encapsulateFieldCollectionsProviderFactory;
        private readonly IEncapsulateFieldConflictFinderFactory _encapsulateFieldConflictFinderFactory;

        public EncapsulateFieldModelFactory(
            IEncapsulateFieldUseBackingUDTMemberModelFactory encapsulateFieldUseBackingUDTMemberModelFactory,
            IEncapsulateFieldUseBackingFieldModelFactory encapsulateFieldUseBackingFieldModelFactory,
            IEncapsulateFieldCollectionsProviderFactory encapsulateFieldCollectionsProviderFactory,
            IEncapsulateFieldConflictFinderFactory encapsulateFieldConflictFinderFactory)
        {
            _useBackingUDTMemberModelFactory = encapsulateFieldUseBackingUDTMemberModelFactory as IEncapsulateFieldUseBackingUDTMemberModelFactory;
            _useBackingFieldModelFactory = encapsulateFieldUseBackingFieldModelFactory;
            _encapsulateFieldCollectionsProviderFactory = encapsulateFieldCollectionsProviderFactory;
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

            var collectionsProvider = _encapsulateFieldCollectionsProviderFactory.Create(targetField.QualifiedModuleName);

            var useBackingFieldModel = _useBackingFieldModelFactory.Create(collectionsProvider, fieldEncapsulationModels);

            var useBackingUDTMemberModel = _useBackingUDTMemberModelFactory.Create(collectionsProvider, fieldEncapsulationModels);

            var initialStrategy = useBackingUDTMemberModel.ObjectStateUDTField.IsExistingDeclaration
                ? EncapsulateFieldStrategy.ConvertFieldsToUDTMembers
                : EncapsulateFieldStrategy.UseBackingFields;

            var conflictFinder = _encapsulateFieldConflictFinderFactory.Create(collectionsProvider);
            var model = new EncapsulateFieldModel(useBackingFieldModel, useBackingUDTMemberModel, conflictFinder)
            {
                EncapsulateFieldStrategy = initialStrategy,
            };

            return model;
        }
    }
}
