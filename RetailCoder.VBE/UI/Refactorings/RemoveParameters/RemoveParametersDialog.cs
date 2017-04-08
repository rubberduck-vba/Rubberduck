using System.Windows.Forms;

namespace Rubberduck.UI.Refactorings.RemoveParameters
{
    public partial class RemoveParametersDialog : Form, IRefactoringDialog<RemoveParametersViewModel>
    {
        public RemoveParametersViewModel ViewModel { get; }

        private RemoveParametersDialog()
        {
            InitializeComponent();
            Text = RubberduckUI.RemoveParamsDialog_Caption;
        }

        public RemoveParametersDialog(RemoveParametersViewModel vm) : this()
        {
            ViewModel = vm;
            RemoveParametersViewElement.DataContext = vm;
            vm.OnWindowClosed += ViewModel_OnWindowClosed;
        }

        void ViewModel_OnWindowClosed(object sender, DialogResult result)
        {
            DialogResult = result;
            Close();
        }
    }
}
