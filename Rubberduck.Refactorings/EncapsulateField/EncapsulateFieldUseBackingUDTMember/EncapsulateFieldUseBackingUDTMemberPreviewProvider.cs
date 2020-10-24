using Rubberduck.Parsing.Rewriter;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.Resources;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.EncapsulateFieldUseBackingUDTMember
{
    public class EncapsulateFieldUseBackingUDTMemberPreviewProvider : RefactoringPreviewProviderWrapperBase<EncapsulateFieldUseBackingUDTMemberModel>
    {
        private readonly INewContentAggregatorFactory _aggregatorFactory;

        public EncapsulateFieldUseBackingUDTMemberPreviewProvider(EncapsulateFieldUseBackingUDTMemberRefactoringAction refactoringAction,
            IRewritingManager rewritingManager,
            INewContentAggregatorFactory aggregatorFactory)
            : base(refactoringAction, rewritingManager)
        {
            _aggregatorFactory = aggregatorFactory;
        }

        public override string Preview(EncapsulateFieldUseBackingUDTMemberModel model)
        {
            model.NewContentAggregator = _aggregatorFactory.Create();
            model.NewContentAggregator.AddNewContent(RubberduckUI.EncapsulateField_PreviewMarker, RubberduckUI.EncapsulateField_PreviewMarker);
            return base.Preview(model);
        }

        protected override QualifiedModuleName ComponentToShow(EncapsulateFieldUseBackingUDTMemberModel model)
        {
            return model.QualifiedModuleName;
        }
    }
}
