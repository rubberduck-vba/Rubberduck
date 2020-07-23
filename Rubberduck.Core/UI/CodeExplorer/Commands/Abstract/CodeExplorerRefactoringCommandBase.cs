using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.UI.Command.Refactorings.Notifiers;
using Rubberduck.VBEditor.Events;

namespace Rubberduck.UI.CodeExplorer.Commands.Abstract
{
    public abstract class CodeExplorerRefactoringCommandBase<TModel> : CodeExplorerCommandBase
        where TModel : class, IRefactoringModel
    {
        private readonly IParserStatusProvider _parserStatusProvider;

        private readonly IRefactoringAction<TModel> _refactoringAction;
        private readonly IRefactoringFailureNotifier _failureNotifier;

        protected CodeExplorerRefactoringCommandBase(
            IRefactoringAction<TModel> refactoringAction,
            IRefactoringFailureNotifier failureNotifier,
            IParserStatusProvider parserStatusProvider,
            IVbeEvents vbeEvents)
            : base(vbeEvents)
        {
            _refactoringAction = refactoringAction;
            _failureNotifier = failureNotifier;

            _parserStatusProvider = parserStatusProvider;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            return _parserStatusProvider.Status == ParserState.Ready;
        }

        protected abstract TModel ModelFromParameter(object parameter);
        protected abstract void ValidateModel(TModel model);

        protected override void OnExecute(object parameter)
        {
            if (!CanExecute(parameter))
            {
                return;
            }

            try
            {
                var model = ModelFromParameter(parameter);
                ValidateModel(model);
                _refactoringAction.Refactor(model);
            }
            catch (RefactoringAbortedException)
            { }
            catch (RefactoringException exception)
            {
                _failureNotifier.Notify(exception);
            }
        }
    }
}