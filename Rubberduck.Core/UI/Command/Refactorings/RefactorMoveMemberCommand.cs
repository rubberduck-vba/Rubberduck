using Rubberduck.Interaction;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.MoveMember;
using Rubberduck.Refactorings.MoveMember.Extensions;
using Rubberduck.UI.Command.Refactorings.Notifiers;
using Rubberduck.VBEditor.Utility;
using System.Runtime.InteropServices;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorMoveMemberCommand : RefactorCodePaneCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly ISelectionProvider _selectionProvider;
        private readonly ISelectedDeclarationProvider _selectedDeclarationProvider;

        public RefactorMoveMemberCommand(
            MoveMemberRefactoring refactoring, 
            MoveMemberFailedNotifier moveMemberFailedNotifier, 
            RubberduckParserState state,
            ISelectionProvider selectionProvider,
            ISelectedDeclarationProvider selectedDeclarationProvider)
            : base(refactoring, moveMemberFailedNotifier, selectionProvider, state)
        {
            _state = state;
            _selectionProvider = selectionProvider;
            _selectedDeclarationProvider = selectedDeclarationProvider;
            AddToCanExecuteEvaluation(SpecializedEvaluateCanExecute);
        }

        private bool SpecializedEvaluateCanExecute(object parameter)
        {
            var target = GetTarget();

            return target != null
                && !_state.IsNewOrModified(target.QualifiedModuleName);
        }

        private Declaration GetTarget()
        {
            var selected = _selectedDeclarationProvider.SelectedDeclaration();
            if (selected is null 
                || !(selected.IsMember()
                        || selected.IsModuleConstant()
                        || (selected.IsField() && !selected.HasPrivateAccessibility())))
            {
                return null;
            }
            return selected;
        }
    }
}
