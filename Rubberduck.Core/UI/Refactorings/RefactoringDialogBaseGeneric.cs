using Rubberduck.Refactorings;

namespace Rubberduck.UI.Refactorings
{
    public class RefactoringDialogBase<TModel, TView, TViewModel> : RefactoringDialogBase, IRefactoringDialog<TModel, TView, TViewModel>
        where TModel : class
        where TView : System.Windows.Controls.UserControl, IRefactoringView, new()
        where TViewModel : class, IRefactoringViewModel<TModel>
    {
        public RefactoringDialogBase(TModel model, TViewModel viewModel) 
        {
            Model = model;
            ViewModel = viewModel;

            View = new TView
            {
                DataContext = ViewModel
            };
            ViewModel.OnWindowClosed += ViewModel_OnWindowClosed;

            UserControl = View;
        }

        public TModel Model { get; }
        public TView View { get; }
        public TViewModel ViewModel { get; }

        public new RefactoringDialogResult DialogResult { get; protected set; }
        public new virtual RefactoringDialogResult ShowDialog()
        {
            var result = base.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK || result == System.Windows.Forms.DialogResult.Yes)
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

        public object DataContext => UserControl.DataContext;
    }
}
