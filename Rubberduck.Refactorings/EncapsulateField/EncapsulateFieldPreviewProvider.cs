using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.EncapsulateFieldUseBackingField;
using Rubberduck.Refactorings.EncapsulateFieldUseBackingUDTMember;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulateFieldPreviewProvider : IRefactoringPreviewProvider<EncapsulateFieldModel>
    {
        private readonly EncapsulateFieldUseBackingFieldPreviewProvider _useBackingFieldPreviewer;
        private readonly EncapsulateFieldUseBackingUDTMemberPreviewProvider _useUDTMembmerPreviewer;
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        public EncapsulateFieldPreviewProvider(IDeclarationFinderProvider declarationFinderProvider,
            EncapsulateFieldUseBackingFieldPreviewProvider useBackingFieldPreviewProvider,
            EncapsulateFieldUseBackingUDTMemberPreviewProvider useUDTMemberPreviewProvide)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _useBackingFieldPreviewer = useBackingFieldPreviewProvider;
            _useUDTMembmerPreviewer = useUDTMemberPreviewProvide;
        }

        public string Preview(EncapsulateFieldModel model)
        {
            var preview = model.EncapsulateFieldStrategy == EncapsulateFieldStrategy.ConvertFieldsToUDTMembers
                ? _useUDTMembmerPreviewer.Preview(model.EncapsulateFieldUseBackingUDTMemberModel)
                : _useBackingFieldPreviewer.Preview(model.EncapsulateFieldUseBackingFieldModel);

            return preview;
        }
    }
}
