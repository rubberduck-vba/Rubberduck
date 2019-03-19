using Rubberduck.Refactorings.Exceptions;

namespace Rubberduck.UI.Command.Refactorings.Notifiers
{
    internal interface IRefactoringFailureNotifier
    {
        void Notify(RefactoringException exception);
    }
}
