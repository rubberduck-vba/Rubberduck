using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.EncapsulateFieldUseBackingField;
using Rubberduck.Refactorings.EncapsulateFieldUseBackingUDTMember;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulateFieldPreviewProvider : IRefactoringPreviewProvider<EncapsulateFieldModel>
    {
        private readonly EncapsulateFieldUseBackingFieldPreviewProvider _useBackingFieldPreviewer;
        private readonly EncapsulateFieldUseBackingUDTMemberPreviewProvider _useBackingUDTMemberPreviewer;
        public EncapsulateFieldPreviewProvider(
            EncapsulateFieldUseBackingFieldPreviewProvider useBackingFieldPreviewProvider,
            EncapsulateFieldUseBackingUDTMemberPreviewProvider useBackingUDTMemberPreviewProvide)
        {
            _useBackingFieldPreviewer = useBackingFieldPreviewProvider;
            _useBackingUDTMemberPreviewer = useBackingUDTMemberPreviewProvide;
        }

        public string Preview(EncapsulateFieldModel model)
        {
            var preview = model.EncapsulateFieldStrategy == EncapsulateFieldStrategy.ConvertFieldsToUDTMembers
                ? _useBackingUDTMemberPreviewer.Preview(model.EncapsulateFieldUseBackingUDTMemberModel)
                : _useBackingFieldPreviewer.Preview(model.EncapsulateFieldUseBackingFieldModel);

            return preview;
        }
    }
}
