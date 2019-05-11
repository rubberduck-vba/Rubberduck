using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.UI.Command.Refactorings.Notifiers;

namespace Rubberduck.UI.Command.Refactorings
{
    public abstract class RefactorCommandBase : CommandBase
    {
        protected readonly IRefactoring Refactoring;
        protected readonly IRefactoringFailureNotifier FailureNotifier;
        protected readonly IParserStatusProvider ParserStatusProvider;

        protected RefactorCommandBase(IRefactoring refactoring, IRefactoringFailureNotifier failureNotifier, IParserStatusProvider parserStatusProvider)
        {
            Refactoring = refactoring;
            ParserStatusProvider = parserStatusProvider;
            FailureNotifier = failureNotifier;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            return ParserStatusProvider.Status == ParserState.Ready;
        }
    }
}