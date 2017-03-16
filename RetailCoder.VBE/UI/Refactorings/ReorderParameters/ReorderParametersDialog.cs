using System.Windows.Forms;

namespace Rubberduck.UI.Refactorings.ReorderParameters
{
    public partial class ReorderParametersDialog : Form, IRefactoringDialog<ReorderParametersViewModel>
    {
        public ReorderParametersViewModel ViewModel { get; }

        private ReorderParametersDialog()
        {
            InitializeComponent();
            Text = RubberduckUI.ReorderParamsDialog_Caption;
        }

        public ReorderParametersDialog(ReorderParametersViewModel vm) : this()
        {
            ViewModel = vm;
            ReorderParametersViewElement.DataContext = vm;
            vm.OnWindowClosed += ViewModel_OnWindowClosed;
        }

        void ViewModel_OnWindowClosed(object sender, DialogResult result)
        {
            DialogResult = result;
            Close();
        }
    }
}
