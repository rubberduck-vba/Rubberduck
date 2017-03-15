using System.Windows.Forms;

namespace Rubberduck.UI.Refactorings.EncapsulateField
{
    public partial class EncapsulateFieldDialog : Form
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
        }

        void ViewModel_OnWindowClosed(object sender, DialogResult result)
        {
            DialogResult = result;
            Close();
        }
    }
}
