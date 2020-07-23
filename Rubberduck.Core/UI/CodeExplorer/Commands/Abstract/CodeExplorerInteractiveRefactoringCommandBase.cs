using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.UI.Command.Refactorings.Notifiers;
using Rubberduck.VBEditor.Events;

namespace Rubberduck.UI.CodeExplorer.Commands.Abstract
{
    public abstract class CodeExplorerInteractiveRefactoringCommandBase<TModel> : CodeExplorerRefactoringCommandBase<TModel>
        where TModel : class, IRefactoringModel
    {
        private readonly IRefactoringAction<TModel> _refactoringAction;
        private readonly IRefactoringUserInteraction<TModel> _refactoringUserInteraction;
        private readonly IRefactoringFailureNotifier _failureNotifier;

        protected CodeExplorerInteractiveRefactoringCommandBase(
            IRefactoringAction<TModel> refactoringAction,
            IRefactoringUserInteraction<TModel> refactoringUserInteraction,
            IRefactoringFailureNotifier failureNotifier,
            IParserStatusProvider parserStatusProvider,
            IVbeEvents vbeEvents)
            : base(refactoringAction, failureNotifier, parserStatusProvider, vbeEvents)
        {
            _refactoringUserInteraction = refactoringUserInteraction;
            _refactoringAction = refactoringAction;
            _failureNotifier = failureNotifier;
        }

        protected abstract TModel InitialModelFromParameter(object parameter);
        protected abstract void ValidateInitialModel(TModel model);

        protected override TModel ModelFromParameter(object parameter)
        {
            var initialModel = InitialModelFromParameter(parameter);
            ValidateInitialModel(initialModel);
            return _refactoringUserInteraction.UserModifiedModel(initialModel);
        }
    }
}