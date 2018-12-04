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
            addRemoveReferencesWindow1.DataContext = viewModel;
        }
    }
}
