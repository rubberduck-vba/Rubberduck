using System;

namespace Rubberduck.Refactorings
{
    public enum RefactoringDialogResult
    {
        Execute,
        Cancel
    }

    public interface IRefactoringDialog<TViewModel> : IRefactoringDialog
    {
        TViewModel ViewModel { get; }
    }

    public interface IRefactoringDialog : IDisposable
    {
        RefactoringDialogResult DialogResult { get; }
        RefactoringDialogResult ShowDialog();
    }
}