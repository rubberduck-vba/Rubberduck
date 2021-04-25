using Rubberduck.Parsing.Rewriter;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.EncapsulateFieldUseBackingField
{
    public class EncapsulateFieldUseBackingFieldPreviewProvider : RefactoringPreviewProviderWrapperBase<EncapsulateFieldUseBackingFieldModel>
    {
        private readonly INewContentAggregatorFactory _aggregatorFactory;

        public EncapsulateFieldUseBackingFieldPreviewProvider(EncapsulateFieldUseBackingFieldRefactoringAction refactoringAction,
            IRewritingManager rewritingManager,
            INewContentAggregatorFactory aggregatorFactory)
            : base(refactoringAction, rewritingManager)
        {
            _aggregatorFactory = aggregatorFactory;
        }

        public override string Preview(EncapsulateFieldUseBackingFieldModel model)
        {
            model.NewContentAggregator = _aggregatorFactory.Create();
            model.NewContentAggregator.AddNewContent(RefactoringsUI.EncapsulateField_PreviewMarker, RefactoringsUI.EncapsulateField_PreviewMarker);
            return base.Preview(model);
        }

        protected override QualifiedModuleName ComponentToShow(EncapsulateFieldUseBackingFieldModel model)
        {
            return model.QualifiedModuleName;
        }
    }
}
