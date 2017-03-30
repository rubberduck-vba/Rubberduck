using System.Windows.Forms;

namespace Rubberduck.UI.Refactorings
{
    public partial class ExtractInterfaceDialog : Form, IRefactoringDialog<ExtractInterfaceViewModel>
    {
        public ExtractInterfaceViewModel ViewModel { get; }

        private ExtractInterfaceDialog()
        {
            InitializeComponent();
            Text = RubberduckUI.ExtractInterface_Caption;
        }

        public ExtractInterfaceDialog(ExtractInterfaceViewModel vm) : this()
        {
            ViewModel = vm;
            ExtractInterfaceViewElement.DataContext = vm;
            vm.OnWindowClosed += ViewModel_OnWindowClosed;
        }

        void ViewModel_OnWindowClosed(object sender, DialogResult result)
        {
            DialogResult = result;
            Close();
        }
    }
}
