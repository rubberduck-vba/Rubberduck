using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.Resources;

namespace Rubberduck.UI.Refactorings.EncapsulateField
{
    public sealed class EncapsulateFieldDialog : RefactoringDialogBase<EncapsulateFieldModel, EncapsulateFieldView, EncapsulateFieldViewModel>
    {
        public EncapsulateFieldDialog(EncapsulateFieldViewModel vm) : base(vm)
        {
            Text = RubberduckUI.EncapsulateField_Caption;
            ViewModel.ExpansionStateChanged += Vm_ExpansionStateChanged;
        }

        private void Vm_ExpansionStateChanged(object sender, bool isExpanded)
        {
            Height = isExpanded ? 560 : 305;
        }
    }
}
