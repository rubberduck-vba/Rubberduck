using Rubberduck.Parsing.Rewriter;
using Rubberduck.VBEditor;
using System;

namespace Rubberduck.Refactorings.EncapsulateFieldUseBackingUDTMember
{
    public class EncapsulateFieldUseBackingUDTMemberPreviewProvider : RefactoringPreviewProviderWrapperBase<EncapsulateFieldUseBackingUDTMemberModel>
    {
        public EncapsulateFieldUseBackingUDTMemberPreviewProvider(EncapsulateFieldUseBackingUDTMemberRefactoringAction refactoringAction,
            IRewritingManager rewritingManager)
            : base(refactoringAction, rewritingManager)
        { }

        public override string Preview(EncapsulateFieldUseBackingUDTMemberModel model)
        {
            var preview = string.Empty;
            var initialFlagValue = model.IncludeNewContentMarker;
            model.IncludeNewContentMarker = true;
            try
            {
                model.ResetNewContent();
                preview = base.Preview(model);
            }
            catch (Exception e) { }
            finally
            {
                model.IncludeNewContentMarker = initialFlagValue;
            }
            return preview;
        }

        protected override QualifiedModuleName ComponentToShow(EncapsulateFieldUseBackingUDTMemberModel model)
        {
            return model.QualifiedModuleName;
        }
    }
}
