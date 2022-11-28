using System.Windows.Forms;
using Rubberduck.UI.Refactorings.ExtractMethod;

namespace Rubberduck.UI.Refactorings
{
    public partial class ExtractMethodDialog : Form, IRefactoringDialog2<ExtractMethodViewModel>
    {
        private ExtractMethodViewModel _viewModel;
        public ExtractMethodViewModel ViewModel
        {
            get => _viewModel;

            set
            {
                _viewModel = value;
                ExtractMethodViewElement.DataContext = ViewModel;
                ViewModel.OnWindowClosed += ViewModel_OnWindowClosed;
            }
        }

        public ExtractMethodDialog()
        {
            InitializeComponent();
            Resize += RepaintHandler;
            Move += RepaintHandler;
        }

        private void RepaintHandler(object sender, System.EventArgs e)
        {
            Invalidate();
            Update();
        }

        public ExtractMethodDialog(ExtractMethodViewModel vm) : this()
        {
            ViewModel = vm;
        }

        private void ViewModel_OnWindowClosed(object sender, DialogResult result)
        {
            DialogResult = result;
            Close();
        }
    }
}
