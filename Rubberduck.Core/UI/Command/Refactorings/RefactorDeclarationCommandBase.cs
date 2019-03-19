using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.UI.Command.Refactorings.Notifiers;

namespace Rubberduck.UI.Command.Refactorings
{
    public abstract class RefactorDeclarationCommandBase : RefactorCommandBase
    {
        protected RefactorDeclarationCommandBase(IRefactoring refactoring, IRefactoringFailureNotifier failureNotifier, IParserStatusProvider parserStatusProvider) 
            : base(refactoring, failureNotifier, parserStatusProvider)
        {}

        protected override void OnExecute(object parameter)
        {
            var target = GetTarget();
            if (target == null)
            {
                return;
            }

            try
            {
                Refactoring.Refactor(target);
            }
            catch (RefactoringAbortedException)
            {
            }
            catch (RefactoringException exception)
            {
                FailureNotifier.Notify(exception);
            }
        }

        protected abstract Declaration GetTarget();
    }
}