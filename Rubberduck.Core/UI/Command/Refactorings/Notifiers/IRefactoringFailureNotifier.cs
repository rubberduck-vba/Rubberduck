using Rubberduck.Refactorings.Exceptions;

namespace Rubberduck.UI.Command.Refactorings.Notifiers
{
    public interface IRefactoringFailureNotifier
    {
        void Notify(RefactoringException exception);
    }
}
