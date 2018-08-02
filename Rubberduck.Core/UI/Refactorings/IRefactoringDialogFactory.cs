using Rubberduck.Refactorings;

namespace Rubberduck.UI.Refactorings
{
    public interface IRefactoringDialogFactory
    {
        TDialog CreateDialog<TModel, TView, TViewModel, TDialog>(TModel model, TView view, TViewModel viewmodel)
            where TModel : class
            where TView : class, IRefactoringView<TModel>
            where TViewModel : class, IRefactoringViewModel<TModel>
            where TDialog : class, IRefactoringDialog<TModel, TView, TViewModel>;
        void ReleaseDialog(IRefactoringDialog dialog);
        TViewModel CreateViewModel<TModel, TViewModel>(TModel model)
            where TModel : class
            where TViewModel : class, IRefactoringViewModel<TModel>;
        void ReleaseViewModel(IRefactoringViewModel viewModel);
    }
}
