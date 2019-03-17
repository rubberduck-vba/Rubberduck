using System;
using NLog;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;

namespace Rubberduck.UI.Command.Refactorings
{
    public abstract class RefactorCommandBase : CommandBase
    {
        protected readonly IRefactoring Refactoring;
        protected readonly IParserStatusProvider ParserStatusProvider;

        protected RefactorCommandBase(IRefactoring refactoring, IParserStatusProvider parserStatusProvider)
            : base(LogManager.GetCurrentClassLogger())
        {
            Refactoring = refactoring;
            ParserStatusProvider = parserStatusProvider;
            CanExecuteEvaluation = StandardEvaluateCanExecute;
        }

        protected Func<object, bool> CanExecuteEvaluation { get; private set; }

        protected void AddToCanExecuteEvaluation(Func<object, bool> furtherCanExecuteEvaluation)
        {
            if (furtherCanExecuteEvaluation == null)
            {
                return;
            }

            CanExecuteEvaluation = (parameter) => CanExecuteEvaluation(parameter) && furtherCanExecuteEvaluation(parameter);
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