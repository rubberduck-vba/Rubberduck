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
        private readonly IEncapsulateFieldRequestFactory _requestFactory;

        public EncapsulateFieldModelFactory(
            IEncapsulateFieldUseBackingUDTMemberModelFactory encapsulateFieldUseBackingUDTMemberModelFactory,
            IEncapsulateFieldUseBackingFieldModelFactory encapsulateFieldUseBackingFieldModelFactory,
            IEncapsulateFieldCandidateCollectionFactory encapsulateFieldCandidateCollectionFactory,
            IEncapsulateFieldRequestFactory encapsulateFieldRequestFactory)
        {
            _useBackingUDTMemberModelFactory = encapsulateFieldUseBackingUDTMemberModelFactory as IEncapsulateFieldUseBackingUDTMemberModelFactory;
            _useBackingFieldModelFactory = encapsulateFieldUseBackingFieldModelFactory;
            _fieldCandidateCollectionFactory = encapsulateFieldCandidateCollectionFactory;
            _requestFactory = encapsulateFieldRequestFactory;
        }

        public EncapsulateFieldModel Create(Declaration target)
        {
            if (!(target is VariableDeclaration targetField))
            {
                throw new ArgumentException();
            }
          
            var fieldCandidates = _fieldCandidateCollectionFactory.Create(targetField.QualifiedModuleName);

            var encapsulationRequest = _requestFactory.Create(targetField);
            var requests = new List<EncapsulateFieldRequest>() { encapsulationRequest };

            var useBackingFieldModel = _useBackingFieldModelFactory.Create(fieldCandidates, requests);

            var useBackingUDTMemberModel = _useBackingUDTMemberModelFactory.Create(fieldCandidates, requests);

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
