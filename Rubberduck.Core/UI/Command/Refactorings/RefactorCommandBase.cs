using System;
using NLog;
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
            : base(LogManager.GetCurrentClassLogger())
        {
            Refactoring = refactoring;
            ParserStatusProvider = parserStatusProvider;
            CanExecuteEvaluation = StandardEvaluateCanExecute;
            FailureNotifier = failureNotifier;
        }

        protected Func<object, bool> CanExecuteEvaluation { get; private set; }

        protected void AddToCanExecuteEvaluation(Func<object, bool> furtherCanExecuteEvaluation)
        {
            if (furtherCanExecuteEvaluation == null)
            {
                return;
            }

            var currentCanExecute = CanExecuteEvaluation; 
            CanExecuteEvaluation = (parameter) => currentCanExecute(parameter) && furtherCanExecuteEvaluation(parameter);
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            return CanExecuteEvaluation(parameter);
        }

        private bool StandardEvaluateCanExecute(object parameter)
        {
            if (ParserStatusProvider.Status != ParserState.Ready)
            {
                return false;
            }

            return true;
        }
    }
}