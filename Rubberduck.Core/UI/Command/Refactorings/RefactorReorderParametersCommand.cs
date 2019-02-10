using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Common;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.ReorderParameters;
using Rubberduck.UI.Refactorings.ReorderParameters;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorReorderParametersCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IRewritingManager _rewritingManager;
        private readonly IMessageBox _msgbox;
        private readonly IRefactoringPresenterFactory _factory;
        private readonly ISelectionService _selectionService;

        public RefactorReorderParametersCommand(IVBE vbe, RubberduckParserState state, IRefactoringPresenterFactory factory, IMessageBox msgbox, IRewritingManager rewritingManager, ISelectionService selectionService) 
            : base (vbe)
        {
            _state = state;
            _rewritingManager = rewritingManager;
            _msgbox = msgbox;
            _factory = factory;
            _selectionService = selectionService;
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

            var activeSelection = _selectionService.ActiveSelection();
            if (!activeSelection.HasValue)
            {
                return false;
            }
            var member = _state.AllUserDeclarations.FindTarget(activeSelection.Value, ValidDeclarationTypes);
            if (member == null || _state.IsNewOrModified(member.QualifiedModuleName))
            {
                return false;
            }

            var parameters = _state.DeclarationFinder.UserDeclarations(DeclarationType.Parameter).Where(item => member.Equals(item.ParentScopeDeclaration)).ToList();
            var canExecute = (member.DeclarationType == DeclarationType.PropertyLet || member.DeclarationType == DeclarationType.PropertySet)
                    ? parameters.Count > 2
                    : parameters.Count > 1;

            return canExecute;
        }

        protected override void OnExecute(object parameter)
        {
            var activeSelection = _selectionService.ActiveSelection();

            if (!activeSelection.HasValue)
            {
                return;
            }

            var refactoring = new ReorderParametersRefactoring(_state, _factory, _msgbox, _rewritingManager, _selectionService);
            refactoring.Refactor(activeSelection.Value);
        }
    }
}
