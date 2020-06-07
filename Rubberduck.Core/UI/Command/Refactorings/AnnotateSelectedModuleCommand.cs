using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.AnnotateDeclaration;
using Rubberduck.UI.Command.Refactorings.Notifiers;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command.Refactorings
{
    public class AnnotateSelectedModuleCommand : AnnotateDeclarationCodePaneCommandBase
    {
        private readonly ISelectedDeclarationProvider _selectedDeclarationProvider;

        public AnnotateSelectedModuleCommand(
            AnnotateDeclarationRefactoring refactoring,
            AnnotateDeclarationFailedNotifier failureNotifier, 
            ISelectionProvider selectionProvider, 
            IParserStatusProvider parserStatusProvider,
            RubberduckParserState state,
            ISelectedDeclarationProvider selectedDeclarationProvider) 
            : base(refactoring, failureNotifier, selectionProvider, parserStatusProvider, state)
        {
            _selectedDeclarationProvider = selectedDeclarationProvider;
        }

        protected override Declaration GetTarget()
        {
            return _selectedDeclarationProvider.SelectedModule();
        }
    }
}