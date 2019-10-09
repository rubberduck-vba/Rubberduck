using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ExtractInterface;
using Rubberduck.UI.Command.Refactorings.Notifiers;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorExtractInterfaceCommand : RefactorCodePaneCommandBase
    {
        private readonly RubberduckParserState _state;

        public RefactorExtractInterfaceCommand(ExtractInterfaceRefactoring refactoring, ExtractInterfaceFailedNotifier extractInterfaceFailedNotifier, RubberduckParserState state, ISelectionService selectionService)
            :base(refactoring, extractInterfaceFailedNotifier, selectionService, state)
        {
            _state = state;

            AddToCanExecuteEvaluation(SpecializedEvaluateCanExecute);
        }

        private bool SpecializedEvaluateCanExecute(object parameter)
        {
            var activeSelection = SelectionService.ActiveSelection();
            if (!activeSelection.HasValue)
            {
                return false;
            }
            return ((ExtractInterfaceRefactoring)Refactoring).CanExecute(_state, activeSelection.Value.QualifiedName);
        }

        private static readonly IReadOnlyList<DeclarationType> ModuleTypes = new[]
        {
            DeclarationType.ClassModule,
            DeclarationType.UserForm,
            DeclarationType.Document,
        };
    }
}
