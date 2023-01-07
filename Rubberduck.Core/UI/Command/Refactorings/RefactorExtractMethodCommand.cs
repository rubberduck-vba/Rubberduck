using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ExtractMethod;
using Rubberduck.UI.Command.Refactorings.Notifiers;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command.Refactorings
{
    public class RefactorExtractMethodCommand : RefactorCodePaneCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly ISelectedDeclarationProvider _selectedDeclarationProvider;
        private readonly ISelectionProvider _selectionProvider;

        public RefactorExtractMethodCommand(
            ExtractMethodRefactoring refactoring, 
            ExtractMethodFailedNotifier failureNotifier, 
            RubberduckParserState state, 
            ISelectionProvider selectionProvider, 
            ISelectedDeclarationProvider selectedDeclarationProvider)
            : base(refactoring, failureNotifier, selectionProvider, state)
        {
            _state = state;
            _selectedDeclarationProvider = selectedDeclarationProvider;
            _selectionProvider = selectionProvider;

            AddToCanExecuteEvaluation(SpecializedEvaluateCanExecute);
        }
        private bool SpecializedEvaluateCanExecute(object parameter)
        {
            var member = _selectedDeclarationProvider.SelectedDeclaration();
            //var moduleContext = _selectedDeclarationProvider.SelectedModule().Context;
            var moduleName = _selectedDeclarationProvider.SelectedModule().QualifiedModuleName;

            if (member == null || _state.IsNewOrModified(member.QualifiedModuleName) || !_selectionProvider.Selection(moduleName).HasValue)
            {
                return false;
            }

            return true;
            //var parameters = _state.DeclarationFinder
            //    .UserDeclarations(DeclarationType.Parameter)
            //    .Where(item => member.Equals(item.ParentScopeDeclaration))
            //    .ToList();

            //return member.DeclarationType == DeclarationType.PropertyLet
            //        || member.DeclarationType == DeclarationType.PropertySet
            //    ? parameters.Count > 2
            //    : parameters.Count > 1;
        }
    }
}
