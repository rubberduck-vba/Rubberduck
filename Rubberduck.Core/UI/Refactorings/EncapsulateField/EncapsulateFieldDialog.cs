using Rubberduck.Refactorings.EncapsulateField;

namespace Rubberduck.UI.Refactorings.EncapsulateField
{
    public sealed class EncapsulateFieldDialog : RefactoringDialogBase<EncapsulateFieldModel, EncapsulateFieldView, EncapsulateFieldViewModel>
    {
        private bool _isExpanded;
        private new int MinHeight => _isExpanded ? 560 : 305;

        public EncapsulateFieldDialog(DialogData dialogData, EncapsulateFieldModel model, EncapsulateFieldView view, EncapsulateFieldViewModel viewModel) : base(dialogData, model, view, viewModel)
        {
            ViewModel.ExpansionStateChanged += Vm_ExpansionStateChanged;
        }

        private void Vm_ExpansionStateChanged(object sender, bool isExpanded)
        {
            _isExpanded = isExpanded;
            Height = MinHeight;
        }
    }
}
