using System;
using NLog;
using Rubberduck.Refactorings;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.Refactorings
{
    public abstract class RefactoringViewModelBase<TModel> : ViewModelBase, IRefactoringViewModel<TModel>
    {
        protected RefactoringViewModelBase(TModel model)
        {
            Model = model;

            OkButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => DialogOk());
            CancelButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => DialogCancel());
        }

        public event EventHandler<RefactoringDialogResult> OnWindowClosed;
        public TModel Model { get; }
        protected virtual void DialogCancel() => OnWindowClosed?.Invoke(this, RefactoringDialogResult.Cancel);
        protected virtual void DialogOk() => OnWindowClosed?.Invoke(this, RefactoringDialogResult.Execute);

        public CommandBase OkButtonCommand { get; }
        public CommandBase CancelButtonCommand { get; }
    }
}
