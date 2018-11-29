using System.Windows.Forms;

namespace Rubberduck.UI.AddRemoveReferences
{
    public partial class AddRemoveReferencesDialog : Form
    {
        public AddRemoveReferencesDialog()
        {
            InitializeComponent();
        }

        public AddRemoveReferencesDialog(AddRemoveReferencesViewModel viewModel) : this()
        {
            addRemoveReferencesWindow1.DataContext = viewModel;
        }
    }
}
