using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.AnnotateDeclaration;
using Rubberduck.UI.Command.Refactorings.Notifiers;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command.Refactorings
{
    public class CodePaneAnnotateDeclarationCommand : RefactorCodePaneCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly ISelectedDeclarationProvider _selectedDeclarationProvider;

        public CodePaneAnnotateDeclarationCommand(
            AnnotateDeclarationRefactoring refactoring,
            AnnotateDeclarationFailedNotifier failureNotifier, 
            ISelectionProvider selectionProvider, 
            IParserStatusProvider parserStatusProvider,
            RubberduckParserState state,
            ISelectedDeclarationProvider selectedDeclarationProvider) 
            : base(refactoring, failureNotifier, selectionProvider, parserStatusProvider)
        {
            _selectedDeclarationProvider = selectedDeclarationProvider;
            _state = state;

            AddToCanExecuteEvaluation(SpecializedEvaluateCanExecute);
        }

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

        private Declaration GetTarget()
        {
            return _selectedDeclarationProvider.SelectedDeclaration();
        }
    }
}