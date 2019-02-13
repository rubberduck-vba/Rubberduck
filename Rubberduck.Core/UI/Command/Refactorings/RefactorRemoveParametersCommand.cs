using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Common;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.RemoveParameters;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorRemoveParametersCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IRefactoringPresenterFactory _factory;

        public RefactorRemoveParametersCommand(RubberduckParserState state, IRefactoringPresenterFactory factory, IRewritingManager rewritingManager, ISelectionService selectionService) 
            : base (rewritingManager, selectionService)
        {
            _state = state;
            _factory = factory;
        }

        private static readonly DeclarationType[] ValidDeclarationTypes =
        {
            DeclarationType.Event,
            DeclarationType.Function,
            DeclarationType.Procedure,
            DeclarationType.PropertyGet,
            DeclarationType.PropertyLet,
            DeclarationType.PropertySet
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

            var member = _state.DeclarationFinder.AllUserDeclarations.FindTarget(activeSelection.Value, ValidDeclarationTypes);
            if (member == null)
            {
                return false;
            }
            if (_state.IsNewOrModified(member.QualifiedModuleName))
            {
                return false;
            }

            var parameters = _state.DeclarationFinder.UserDeclarations(DeclarationType.Parameter)
                .Where(item => member.Equals(item.ParentScopeDeclaration))
                .ToList();
            return member.DeclarationType == DeclarationType.PropertyLet 
                    || member.DeclarationType == DeclarationType.PropertySet
                        ? parameters.Count > 1
                        : parameters.Any();
        }

        protected override void OnExecute(object parameter)
        {
            var activeSelection = SelectionService.ActiveSelection();
            if (!activeSelection.HasValue)
            {
                return;
            }

            var refactoring = new RemoveParametersRefactoring(_state, _factory, RewritingManager, SelectionService);
            refactoring.Refactor(activeSelection.Value);
        }
    }
}
