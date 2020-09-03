using Rubberduck.Parsing.Rewriter;
using Rubberduck.VBEditor;
using System;

namespace Rubberduck.Refactorings.EncapsulateFieldUseBackingField
{
    public class EncapsulateFieldUseBackingFieldPreviewProvider : RefactoringPreviewProviderWrapperBase<EncapsulateFieldUseBackingFieldModel>
    {
        public EncapsulateFieldUseBackingFieldPreviewProvider(EncapsulateFieldUseBackingFieldRefactoringAction refactoringAction,
            IRewritingManager rewritingManager)
            : base(refactoringAction, rewritingManager)
        { }

        public override string Preview(EncapsulateFieldUseBackingFieldModel model)
        {
            var preview = string.Empty;
            var initialFlagValue = model.IncludeNewContentMarker;
            model.IncludeNewContentMarker = true;
            try
            {
                model.ResetNewContent();
                preview = base.Preview(model);
            }
            catch (Exception) { }
            finally
            {
                model.IncludeNewContentMarker = initialFlagValue;
            }
            return preview;
        }

        protected override QualifiedModuleName ComponentToShow(EncapsulateFieldUseBackingFieldModel model)
        {
            return model.QualifiedModuleName;
        }
    }
}
