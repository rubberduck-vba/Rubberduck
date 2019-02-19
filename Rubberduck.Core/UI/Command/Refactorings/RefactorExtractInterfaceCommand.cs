using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Antlr4.Runtime;
using Rubberduck.Interaction;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.ExtractInterface;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorExtractInterfaceCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _messageBox;
        private readonly IRefactoringPresenterFactory _factory;
        
        public RefactorExtractInterfaceCommand(RubberduckParserState state, IMessageBox messageBox, IRefactoringPresenterFactory factory, IRewritingManager rewritingManager, ISelectionService selectionService)
            :base(rewritingManager, selectionService)
        {
            _state = state;
            _messageBox = messageBox;
            _factory = factory;
        }

        private static readonly IReadOnlyList<DeclarationType> ModuleTypes = new[] 
        {
            DeclarationType.ClassModule,
            DeclarationType.UserForm, 
            DeclarationType.ProceduralModule, 
        };

        protected override bool EvaluateCanExecute(object parameter)
        {
            if (_state.Status != ParserState.Ready)
            {
                return false;
            }

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

        protected override void OnExecute(object parameter)
        {
            var activeSelection = SelectionService.ActiveSelection();
            if (!activeSelection.HasValue)
            {
                return;
            }

            var refactoring = new ExtractInterfaceRefactoring(_state, _state, _messageBox, _factory, RewritingManager, SelectionService);
            refactoring.Refactor(activeSelection.Value);
        }
    }
}
