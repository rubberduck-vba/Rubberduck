using System.Windows.Forms;

namespace Rubberduck.UI.Refactorings
{
    public interface IRefactoringDialog<T>
    {
        T ViewModel { get; }
        DialogResult DialogResult { get; }

        DialogResult ShowDialog();
    }
}