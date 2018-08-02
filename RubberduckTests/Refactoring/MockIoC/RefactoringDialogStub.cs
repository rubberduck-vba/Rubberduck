using System.Diagnostics.CodeAnalysis;
using Rubberduck.Refactorings;

namespace RubberduckTests.Refactoring.MockIoC
{
    internal class RefactoringDialogStub<TModel, TView, TViewModel> : IRefactoringDialog<TModel, TView, TViewModel>
        where TModel : class
        where TView : class, IRefactoringView<TModel>
        where TViewModel : class, IRefactoringViewModel<TModel>
    {
        [SuppressMessage("ReSharper", "VirtualMemberCallInConstructor")]
        public RefactoringDialogStub(TModel model, TView view, TViewModel viewModel)
        {
            Model = model;
            ViewModel = viewModel;

            View = view;
            View.DataContext = viewModel;
            ViewModel.OnWindowClosed += ViewModel_OnWindowClosed;

            DialogResult = RefactoringDialogResult.Execute;
        }

        public virtual void Dispose()
        {
            
        }

        public virtual RefactoringDialogResult DialogResult { get; }
        public virtual RefactoringDialogResult ShowDialog()
        {
            return DialogResult;
        }

        public virtual TModel Model { get; }
        public virtual TView View { get; }
        public virtual TViewModel ViewModel { get; }

        protected virtual void ViewModel_OnWindowClosed(object sender, RefactoringDialogResult result)
        {

        }
    }
}
