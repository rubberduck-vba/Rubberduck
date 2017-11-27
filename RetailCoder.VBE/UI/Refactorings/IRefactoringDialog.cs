using System;
using System.Windows.Forms;

namespace Rubberduck.UI.Refactorings
{
    public interface IRefactoringDialog<TViewModel> : IDisposable
    {
        TViewModel ViewModel { get; }
        DialogResult DialogResult { get; }

        DialogResult ShowDialog();
    }

    // TODO: this should be consolidated....
    public interface IRefactoringDialog2<TViewModel> : IRefactoringDialog<TViewModel>
    {
        new TViewModel ViewModel { get; set; }
    }
}