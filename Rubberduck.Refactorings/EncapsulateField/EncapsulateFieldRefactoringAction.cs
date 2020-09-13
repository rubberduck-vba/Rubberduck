using Rubberduck.Refactorings.EncapsulateFieldUseBackingField;
using Rubberduck.Refactorings.EncapsulateFieldUseBackingUDTMember;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulateFieldRefactoringAction : IRefactoringAction<EncapsulateFieldModel>
    {
        private readonly EncapsulateFieldUseBackingFieldRefactoringAction _useBackingField;
        private readonly EncapsulateFieldUseBackingUDTMemberRefactoringAction _useBackingUDTMember;
        private readonly INewContentAggregatorFactory _newContentAggregatorFactory;

        public EncapsulateFieldRefactoringAction(
            EncapsulateFieldUseBackingFieldRefactoringAction encapsulateFieldUseBackingField,
            EncapsulateFieldUseBackingUDTMemberRefactoringAction encapsulateFieldUseUDTMember,
            INewContentAggregatorFactory newContentAggregatorFactory)
        {
            _useBackingField = encapsulateFieldUseBackingField;
            _useBackingUDTMember = encapsulateFieldUseUDTMember;
            _newContentAggregatorFactory = newContentAggregatorFactory;
        }

        public void Refactor(EncapsulateFieldModel model)
        {
            if (!model?.EncapsulationCandidates.Any() ?? true)
            {
                return;
            }

            if (model.EncapsulateFieldStrategy == EncapsulateFieldStrategy.ConvertFieldsToUDTMembers)
            {
                model.EncapsulateFieldUseBackingUDTMemberModel.NewContentAggregator = _newContentAggregatorFactory.Create();
                _useBackingUDTMember.Refactor(model.EncapsulateFieldUseBackingUDTMemberModel);
                return;
            }

            model.EncapsulateFieldUseBackingFieldModel.NewContentAggregator = _newContentAggregatorFactory.Create();
            _useBackingField.Refactor(model.EncapsulateFieldUseBackingFieldModel);
        }
    }
}
