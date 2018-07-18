using Rubberduck.Refactorings;

namespace RubberduckTests.Refactoring.MockIoC
{
    internal class RefactoringDialogStub<TModel, TView, TViewModel> : IRefactoringDialog<TModel, TView, TViewModel>
        where TModel : class
        where TView : class, IRefactoringView<TModel>
        where TViewModel : class, IRefactoringViewModel<TModel>
    {
        public RefactoringDialogStub(TModel model, TView view, TViewModel viewModel)
        {
            Model = model;
            ViewModel = viewModel;

            View = view;
            View.DataContext = viewModel;
            ViewModel.OnWindowClosed += ViewModel_OnWindowClosed;
        }

public void Dispose()
        {
            //no-op
        }

        public RefactoringDialogResult DialogResult { get; protected set; }
        public RefactoringDialogResult ShowDialog()
        {
            return DialogResult;
        }

        public TModel Model { get; }
        public TView View { get; }
        public TViewModel ViewModel { get; }

        protected virtual void ViewModel_OnWindowClosed(object sender, RefactoringDialogResult result)
        {
            DialogResult = result;
        }
    }
}
