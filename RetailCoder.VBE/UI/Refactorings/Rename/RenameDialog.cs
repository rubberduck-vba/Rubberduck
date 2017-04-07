using System.Windows.Forms;

namespace Rubberduck.UI.Refactorings.Rename
{
    public partial class RenameDialog : Form, IRefactoringDialog<RenameViewModel>
    {
        public RenameViewModel ViewModel { get; }

        private RenameDialog()
        {
            InitializeComponent();
            Text = RubberduckUI.RenameDialog_Caption;
        }

        public RenameDialog(RenameViewModel vm) : this()
        {
            ViewModel = vm;
            RenameViewElement.DataContext = vm;
            vm.OnWindowClosed += ViewModel_OnWindowClosed;
        }

        void ViewModel_OnWindowClosed(object sender, DialogResult result)
        {
            DialogResult = result;
            Close();
        }
    }
}
