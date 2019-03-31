using Rubberduck.Refactorings;

namespace Rubberduck.UI.Refactorings
{
    public interface IRefactoringPresenter<out TModel, out TDialog, TView, out TViewModel> 
        where TModel : class
        where TView : class, IRefactoringView<TModel>
        where TDialog : class, IRefactoringDialog<TModel, TView, TViewModel> 
        where TViewModel : class, IRefactoringViewModel<TModel>
    {
        TModel Model { get; }
        TDialog Dialog { get; }
        TViewModel ViewModel { get; }
        RefactoringDialogResult DialogResult { get; }
        TModel Show();
    }
}