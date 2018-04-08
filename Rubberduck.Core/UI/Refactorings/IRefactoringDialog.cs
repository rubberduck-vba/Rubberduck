using System;
using System.Windows.Forms;

namespace Rubberduck.UI.Refactorings
{
    public interface IRefactoringDialog<T> : IDisposable
    {
        T ViewModel { get; }
        DialogResult DialogResult { get; }

        DialogResult ShowDialog();
    }
}