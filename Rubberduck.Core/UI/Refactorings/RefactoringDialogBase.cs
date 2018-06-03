using Forms = System.Windows.Forms;
using Rubberduck.Refactorings;

namespace Rubberduck.UI.Refactorings
{
    public partial class RefactoringDialogBase<TModel, TView, TViewModel> : Forms.Form, IRefactoringDialog<TViewModel>
        where TView : System.Windows.Controls.UserControl, new()
        where TViewModel : RefactoringViewModelBase<TModel>
    {
        public RefactoringDialogBase(TViewModel viewModel)
        {
            ViewModel = viewModel;
            userControl = new TView
            {
                DataContext = ViewModel
            };
            ViewModel.OnWindowClosed += ViewModel_OnWindowClosed;
        }

        public TViewModel ViewModel { get; set; }
        public new RefactoringDialogResult DialogResult { get; protected set; }
        public new virtual RefactoringDialogResult ShowDialog()
        {
            var result = base.ShowDialog();
            if (result == Forms.DialogResult.OK || result == Forms.DialogResult.Yes)
            {
                DialogResult = RefactoringDialogResult.Execute;
            }
            else
            {
                DialogResult = RefactoringDialogResult.Cancel;
            }

            return DialogResult;
        }

        protected virtual void ViewModel_OnWindowClosed(object sender, RefactoringDialogResult result)
        {
            DialogResult = result;
            Close();
        }
    }
}
