using System;

namespace Rubberduck.Refactorings
{
    public enum RefactoringDialogResult
    {
        Execute,
        Cancel
    }

    public interface IRefactoringDialog<TModel, TView, TViewModel> : IRefactoringDialog
        where TModel : class
        where TView : class, new()
        where TViewModel : class, IRefactoringViewModel<TModel>
    {
        TViewModel ViewModel { get; }
    }

    public interface IRefactoringDialog : IDisposable
    {
        RefactoringDialogResult DialogResult { get; }
        RefactoringDialogResult ShowDialog();
    }
}