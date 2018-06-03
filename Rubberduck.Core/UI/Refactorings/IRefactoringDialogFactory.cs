namespace Rubberduck.UI.Refactorings
{
    public interface IRefactoringDialogFactory<TModel, TView, TViewModel, TDialog>
        where TModel : class
        where TView : System.Windows.Controls.UserControl, new()
        where TViewModel : RefactoringViewModelBase<TModel>
        where TDialog : RefactoringDialogBase<TModel, TView, TViewModel>
    {
        TDialog CreateDialog(TViewModel viewmodel);
        void ReleaseDialog(TDialog dialog);
        TViewModel CreateViewModel(TModel model);
        void ReleaseViewModel(TViewModel viewModel);
    }

    public interface IRefactoringPresenterFactory<TModel, TPresenter>
    {
        TPresenter Create(TModel model);
        void Release(TPresenter presenter);
    }
}
