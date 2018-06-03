using System;

namespace Rubberduck.Refactorings
{
    public enum RefactoringDialogResult
    {
        Execute,
        Cancel
    }

    public interface IRefactoringDialog<TViewModel> : IDisposable
    {
        TViewModel ViewModel { get; }
        RefactoringDialogResult DialogResult { get; }
        RefactoringDialogResult ShowDialog();
    }
}