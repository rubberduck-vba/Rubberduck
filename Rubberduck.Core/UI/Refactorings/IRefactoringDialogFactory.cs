using Rubberduck.Refactorings;

namespace Rubberduck.UI.Refactorings
{
    public interface IRefactoringDialogFactory
    {
        TView CreateView<TModel, TView>(TModel model)
            where TModel: class
            where TView:class, IRefactoringView<TModel>;
        void ReleaseView(IRefactoringView view);

        TViewModel CreateViewModel<TModel, TViewModel>(TModel model)
            where TModel : class
            where TViewModel : class, IRefactoringViewModel<TModel>;
        void ReleaseViewModel(IRefactoringViewModel viewModel);

        TDialog CreateDialog<TModel, TView, TViewModel, TDialog>(DialogData dialogData, TModel model, TView view, TViewModel viewmodel)
            where TModel : class
            where TView : class, IRefactoringView<TModel>
            where TViewModel : class, IRefactoringViewModel<TModel>
            where TDialog : class, IRefactoringDialog<TModel, TView, TViewModel>;
        void ReleaseDialog(IRefactoringDialog dialog);
    }
}
