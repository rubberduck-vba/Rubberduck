using Rubberduck.Refactorings;

namespace Rubberduck.UI.Refactorings
{
    public interface IRefactoringDialogFactory
    {
        TDialog CreateDialog<TModel, TView, TViewModel, TDialog>(TViewModel viewmodel)
            where TModel : class
            where TView : System.Windows.Controls.UserControl, new()
            where TViewModel : RefactoringViewModelBase<TModel>
            where TDialog : RefactoringDialogBase<TModel, TView, TViewModel>;
        void ReleaseDialog(IRefactoringDialog dialog);
        TViewModel CreateViewModel<TModel, TViewModel>(TModel model)
            where TModel : class
            where TViewModel : RefactoringViewModelBase<TModel>;
        void ReleaseViewModel(IRefactoringViewModel viewModel);
    }
}
