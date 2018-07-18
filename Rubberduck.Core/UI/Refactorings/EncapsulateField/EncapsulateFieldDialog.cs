using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.Resources;

namespace Rubberduck.UI.Refactorings.EncapsulateField
{
    public sealed class EncapsulateFieldDialog : RefactoringDialogBase<EncapsulateFieldModel, EncapsulateFieldView, EncapsulateFieldViewModel>
    {
        private bool _isExpanded;
        protected override int MinWidth => 667;
        protected override int MinHeight => _isExpanded ? 560 : 305;

        public EncapsulateFieldDialog(EncapsulateFieldModel model, EncapsulateFieldView view, EncapsulateFieldViewModel viewModel) : base(model, view, viewModel)
        {
            Text = RubberduckUI.EncapsulateField_Caption;
            ViewModel.ExpansionStateChanged += Vm_ExpansionStateChanged;
        }

        private void Vm_ExpansionStateChanged(object sender, bool isExpanded)
        {
            _isExpanded = isExpanded;
            Height = MinHeight;
        }
    }
}
