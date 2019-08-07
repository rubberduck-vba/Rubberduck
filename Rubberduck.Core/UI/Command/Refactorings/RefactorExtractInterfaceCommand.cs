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

            var interfaceClass = _state.AllUserDeclarations.SingleOrDefault(item =>
                item.QualifiedName.QualifiedModuleName.Equals(activeSelection.Value.QualifiedName)
                && ModuleTypes.Contains(item.DeclarationType));

            if (interfaceClass == null)
            {
                return false;
            }

            // interface class must have members to be implementable
            var hasMembers = _state.AllUserDeclarations.Any(item =>
                item.DeclarationType.HasFlag(DeclarationType.Member)
                && item.ParentDeclaration != null
                && item.ParentDeclaration.Equals(interfaceClass));

            if (!hasMembers)
            {
                return false;
            }

            var parseTree = _state.GetParseTree(interfaceClass.QualifiedName.QualifiedModuleName);
            var context = ((ParserRuleContext)parseTree).GetDescendents<VBAParser.ImplementsStmtContext>();

            // true if active code pane is for a class/document/form module
            return !context.Any()
                   && !_state.IsNewOrModified(interfaceClass.QualifiedModuleName)
                   && !_state.IsNewOrModified(activeSelection.Value.QualifiedName);
        }

        private static readonly IReadOnlyList<DeclarationType> ModuleTypes = new[] 
        {
            DeclarationType.ClassModule,
            DeclarationType.UserForm, 
            DeclarationType.Document, 
        };
    }
}
