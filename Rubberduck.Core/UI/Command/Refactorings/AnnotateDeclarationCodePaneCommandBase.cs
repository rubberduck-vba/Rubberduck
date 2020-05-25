using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.AnnotateDeclaration;
using Rubberduck.UI.Command.Refactorings.Notifiers;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command.Refactorings
{
    public abstract class AnnotateDeclarationCodePaneCommandBase : RefactorCodePaneCommandBase
    {
        private readonly RubberduckParserState _state;

        protected AnnotateDeclarationCodePaneCommandBase(
            AnnotateDeclarationRefactoring refactoring,
            AnnotateDeclarationFailedNotifier failureNotifier, 
            ISelectionProvider selectionProvider, 
            IParserStatusProvider parserStatusProvider,
            RubberduckParserState state) 
            : base(refactoring, failureNotifier, selectionProvider, parserStatusProvider)
        {
            _state = state;

            AddToCanExecuteEvaluation(SpecializedEvaluateCanExecute);
        }

        protected abstract Declaration GetTarget();

        private bool SpecializedEvaluateCanExecute(object parameter)
        {
            var target = GetTarget();

            if (target == null)
            {
                return false;
            }

            var targetType = target.DeclarationType;

            if (!targetType.HasFlag(DeclarationType.Member)
                && !targetType.HasFlag(DeclarationType.Module)
                && !targetType.HasFlag(DeclarationType.Variable)
                && !targetType.HasFlag(DeclarationType.Constant))
            {
                return false;
            }

            return !_state.IsNewOrModified(target.QualifiedModuleName);
        }
    }
}