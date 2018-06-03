using System;
using Rubberduck.Refactorings;

namespace Rubberduck.UI.Refactorings
{
    public class RefactoringViewModelBase<TModel> : ViewModelBase, IRefactoringViewModel<TModel>
    {
        public RefactoringViewModelBase(TModel model)
        {
            Model = model;
        }

        public event EventHandler<RefactoringDialogResult> OnWindowClosed;
        public TModel Model { get; }
        protected virtual void DialogCancel() => OnWindowClosed?.Invoke(this, RefactoringDialogResult.Cancel);
        protected virtual void DialogOk() => OnWindowClosed?.Invoke(this, RefactoringDialogResult.Execute);
    }
}
