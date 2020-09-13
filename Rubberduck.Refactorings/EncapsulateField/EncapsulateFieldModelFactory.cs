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
        private readonly IEncapsulateFieldUseBackingUDTMemberModelFactory _useBackingUDTMemberModelFactory;
        private readonly IEncapsulateFieldUseBackingFieldModelFactory _useBackingFieldModelFactory;
        private readonly IEncapsulateFieldCandidateCollectionFactory _fieldCandidateCollectionFactory;

        public EncapsulateFieldModelFactory(
            IEncapsulateFieldUseBackingUDTMemberModelFactory encapsulateFieldUseBackingUDTMemberModelFactory,
            IEncapsulateFieldUseBackingFieldModelFactory encapsulateFieldUseBackingFieldModelFactory,
            IEncapsulateFieldCandidateCollectionFactory encapsulateFieldCandidateCollectionFactory)
        {
            _useBackingUDTMemberModelFactory = encapsulateFieldUseBackingUDTMemberModelFactory as IEncapsulateFieldUseBackingUDTMemberModelFactory;
            _useBackingFieldModelFactory = encapsulateFieldUseBackingFieldModelFactory;
            _fieldCandidateCollectionFactory = encapsulateFieldCandidateCollectionFactory;
        }

        public EncapsulateFieldModel Create(Declaration target)
        {
            if (!(target is VariableDeclaration targetField))
            {
                throw new ArgumentException();
            }
          
            var fieldCandidates = _fieldCandidateCollectionFactory.Create(targetField.QualifiedModuleName);

            var fieldEncapsulationModels = new List<FieldEncapsulationModel>() { new FieldEncapsulationModel(targetField) };

            var useBackingFieldModel = _useBackingFieldModelFactory.Create(fieldCandidates, fieldEncapsulationModels);

            var useBackingUDTMemberModel = _useBackingUDTMemberModelFactory.Create(fieldCandidates, fieldEncapsulationModels);

            var initialStrategy = useBackingUDTMemberModel.ObjectStateUDTField.IsExistingDeclaration
                ? EncapsulateFieldStrategy.ConvertFieldsToUDTMembers
                : EncapsulateFieldStrategy.UseBackingFields;

            return new EncapsulateFieldModel(useBackingFieldModel, useBackingUDTMemberModel)
            {
                EncapsulateFieldStrategy = initialStrategy,
            };
        }
    }
}
