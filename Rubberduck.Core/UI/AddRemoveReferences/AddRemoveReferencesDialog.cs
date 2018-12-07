using System.Windows.Forms;
using Rubberduck.UI.Refactorings;

namespace Rubberduck.UI.AddRemoveReferences
{
    public partial class AddRemoveReferencesDialog : Form, IRefactoringDialog<AddRemoveReferencesViewModel>
    {
        public AddRemoveReferencesViewModel ViewModel { get; }

        public AddRemoveReferencesDialog()
        {
            InitializeComponent();           
        }

        public AddRemoveReferencesDialog(AddRemoveReferencesViewModel viewModel) : this()
        {
            ViewModel = viewModel;
            ViewModel.OnWindowClosed += ViewModel_OnWindowClosed;
            addRemoveReferencesWindow1.DataContext = viewModel;
        }

        private void ViewModel_OnWindowClosed(object sender, DialogResult result)
        {
            DialogResult = result;
            Close();
        }
    }
}
