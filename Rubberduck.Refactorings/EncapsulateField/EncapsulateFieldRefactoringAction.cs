using Rubberduck.Refactorings.EncapsulateFieldUseBackingField;
using Rubberduck.Refactorings.EncapsulateFieldUseBackingUDTMember;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulateFieldRefactoringAction : IRefactoringAction<EncapsulateFieldModel>
    {
        private readonly EncapsulateFieldUseBackingFieldRefactoringAction _useBackingField;
        private readonly EncapsulateFieldUseBackingUDTMemberRefactoringAction _useBackingUDTMember;

        public EncapsulateFieldRefactoringAction(
            EncapsulateFieldUseBackingFieldRefactoringAction encapsulateFieldUseBackingField,
            EncapsulateFieldUseBackingUDTMemberRefactoringAction encapsulateFieldUseUDTMember)
        {
            _useBackingField = encapsulateFieldUseBackingField;
            _useBackingUDTMember = encapsulateFieldUseUDTMember;
        }

        public void Refactor(EncapsulateFieldModel model)
        {
            if (!model?.EncapsulationCandidates.Any() ?? true)
            {
                return;
            }

            if (model.EncapsulateFieldStrategy == EncapsulateFieldStrategy.ConvertFieldsToUDTMembers)
            {
                _useBackingUDTMember.Refactor(model.EncapsulateFieldUseBackingUDTMemberModel);
                return;
            }

            _useBackingField.Refactor(model.EncapsulateFieldUseBackingFieldModel);
        }
    }
}
