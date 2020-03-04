using System;

namespace Rubberduck.Refactorings
{
    public enum RefactoringDialogResult
    {
        Undetermined,
        Execute,
        Cancel
    }

    public interface IRefactoringDialog<out TModel, out TView, out TViewModel> : IRefactoringDialog
        where TModel : class
        where TView : class, IRefactoringView<TModel>
        where TViewModel : class, IRefactoringViewModel<TModel>
    {
        TModel Model { get; }
        TView View { get; }
        TViewModel ViewModel { get; }
    }

    public interface IRefactoringDialog : IDisposable
    {
        RefactoringDialogResult DialogResult { get; }
        RefactoringDialogResult ShowDialog();
    }
}