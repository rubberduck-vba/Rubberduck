using Rubberduck.Refactorings;

namespace Rubberduck.UI.Refactorings
{
    public class RefactoringDialogBase<TModel, TView, TViewModel> : RefactoringDialogBase, IRefactoringDialog<TModel, TView, TViewModel>
        where TModel : class
        where TView : System.Windows.Controls.UserControl, IRefactoringView<TModel>
        where TViewModel : class, IRefactoringViewModel<TModel>
    {
        public RefactoringDialogBase(TModel model, TView view, TViewModel viewModel) 
        {
            Model = model;
            ViewModel = viewModel;

            View = view;
            View.DataContext = ViewModel;
            ViewModel.OnWindowClosed += ViewModel_OnWindowClosed;

            UserControl = View;
        }

        public TModel Model { get; }
        public TView View { get; }
        public TViewModel ViewModel { get; }
        
        public new RefactoringDialogResult DialogResult { get; protected set; }
        public new virtual RefactoringDialogResult ShowDialog()
        {
            // The return of ShowDialog is meaningless; we use the DialogResult which the commands set.
            var result = base.ShowDialog();
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
