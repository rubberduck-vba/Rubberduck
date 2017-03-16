using System.Windows.Forms;

namespace Rubberduck.UI.Refactorings.EncapsulateField
{
    public partial class EncapsulateFieldDialog : Form, IRefactoringDialog<EncapsulateFieldViewModel>
    {
        public EncapsulateFieldViewModel ViewModel { get; }

        private EncapsulateFieldDialog()
        {
            InitializeComponent();
            Text = RubberduckUI.EncapsulateField_Caption;
        }

        public EncapsulateFieldDialog(EncapsulateFieldViewModel vm) : this()
        {
            ViewModel = vm;
            EncapsulateFieldViewElement.DataContext = vm;
            vm.OnWindowClosed += ViewModel_OnWindowClosed;
            vm.ExpansionStateChanged += Vm_ExpansionStateChanged;
        }

        private void Vm_ExpansionStateChanged(object sender, bool isExpanded)
        {
            Height = isExpanded ? 560 : 305;
        }

        void ViewModel_OnWindowClosed(object sender, DialogResult result)
        {
            DialogResult = result;
            Close();
        }
    }
}
