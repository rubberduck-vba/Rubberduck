using System.Drawing;
using System.Windows.Forms;
using Rubberduck.UI.Refactorings;

namespace Rubberduck.UI.AddRemoveReferences
{
    public partial class AddRemoveReferencesDialog : Form
    {
        public AddRemoveReferencesViewModel ViewModel { get; }

        public AddRemoveReferencesDialog()
        {
            InitializeComponent();
            MinimumSize= new Size(600, 380);
        }

        public sealed override Size MinimumSize
        {
            get => base.MinimumSize;
            set => base.MinimumSize = value;
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
