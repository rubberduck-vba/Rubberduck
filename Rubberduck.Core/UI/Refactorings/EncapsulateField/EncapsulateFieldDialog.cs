using Rubberduck.Refactorings.EncapsulateField;

namespace Rubberduck.UI.Refactorings.EncapsulateField
{
    public sealed class EncapsulateFieldDialog : RefactoringDialogBase<EncapsulateFieldModel, EncapsulateFieldView, EncapsulateFieldViewModel>
    {
        public EncapsulateFieldDialog(DialogData dialogData, EncapsulateFieldModel model, EncapsulateFieldView view, EncapsulateFieldViewModel viewModel) : base(dialogData, model, view, viewModel)
        {
        }
    }
}
