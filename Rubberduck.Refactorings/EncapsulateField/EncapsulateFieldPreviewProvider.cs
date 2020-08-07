using Rubberduck.Parsing.Rewriter;
using Rubberduck.Refactorings.EncapsulateField.Extensions;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulateFieldPreviewProvider : IRefactoringPreviewProvider<EncapsulateFieldModel>
    {
        private readonly EncapsulateFieldUseBackingFieldPreviewProvider _useBackingFieldPreviewer;
        private readonly EncapsulateFieldUseBackingUDTMemberPreviewProvider _useUDTMembmerPreviewer;
        public EncapsulateFieldPreviewProvider(
            EncapsulateFieldUseBackingFieldPreviewProvider useBackingFieldPreviewProvider,
            EncapsulateFieldUseBackingUDTMemberPreviewProvider useUDTMemberPreviewProvide)
        {
            _useBackingFieldPreviewer = useBackingFieldPreviewProvider;
            _useUDTMembmerPreviewer = useUDTMemberPreviewProvide;
        }

        public string Preview(EncapsulateFieldModel model)
        {
            var preview = model.EncapsulateFieldStrategy == EncapsulateFieldStrategy.ConvertFieldsToUDTMembers
                                    ? _useUDTMembmerPreviewer.Preview(model)
                                    : _useBackingFieldPreviewer.Preview(model);

            return preview.LimitNewLines(3);
        }
    }

    public class EncapsulateFieldUseBackingFieldPreviewProvider : RefactoringPreviewProviderWrapperBase<EncapsulateFieldModel>
    {
        public EncapsulateFieldUseBackingFieldPreviewProvider(EncapsulateFieldUseBackingFieldRefactoringAction refactoringAction,
            IRewritingManager rewritingManager)
            : base(refactoringAction, rewritingManager)
        {

        }

        protected override QualifiedModuleName ComponentToShow(EncapsulateFieldModel model)
        {
            return model.QualifiedModuleName;
        }
    }

    public class EncapsulateFieldUseBackingUDTMemberPreviewProvider : RefactoringPreviewProviderWrapperBase<EncapsulateFieldModel>
    {
        public EncapsulateFieldUseBackingUDTMemberPreviewProvider(EncapsulateFieldUseBackingUDTMemberRefactoringAction refactoringAction,
            IRewritingManager rewritingManager)
            : base(refactoringAction, rewritingManager)
        {

        }

        protected override QualifiedModuleName ComponentToShow(EncapsulateFieldModel model)
        {
            return model.QualifiedModuleName;
        }
    }
}
