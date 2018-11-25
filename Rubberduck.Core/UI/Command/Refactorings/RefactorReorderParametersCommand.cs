using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Common;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ReorderParameters;
using Rubberduck.UI.Refactorings.ReorderParameters;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorReorderParametersCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IRewritingManager _rewritingManager;
        private readonly IMessageBox _msgbox;

        public RefactorReorderParametersCommand(IVBE vbe, RubberduckParserState state, IMessageBox msgbox, IRewritingManager rewritingManager) 
            : base (vbe)
        {
            _state = state;
            _rewritingManager = rewritingManager;
            _msgbox = msgbox;
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

            var selection = Vbe.GetActiveSelection();
            if (selection == null)
            {
                return false;
            }
            var member = _state.AllUserDeclarations.FindTarget(selection.Value, ValidDeclarationTypes);
            if (member == null)
            {
                return false;
            }
            if (_state.IsNewOrModified(member.QualifiedModuleName))
            {
                return false;
            }

            var parameters = _state.AllUserDeclarations.Where(item => item.DeclarationType == DeclarationType.Parameter && member.Equals(item.ParentScopeDeclaration)).ToList();
            var canExecute = (member.DeclarationType == DeclarationType.PropertyLet || member.DeclarationType == DeclarationType.PropertySet)
                    ? parameters.Count > 2
                    : parameters.Count > 1;

            return canExecute;
        }

        protected override void OnExecute(object parameter)
        {
            var selection = Vbe.GetActiveSelection();

            if (selection == null)
            {
                return;
            }

            using (var view = new ReorderParametersDialog(new ReorderParametersViewModel(_state)))
            {
                var factory = new ReorderParametersPresenterFactory(Vbe, view, _state, _msgbox);
                var refactoring = new ReorderParametersRefactoring(Vbe, factory, _msgbox, _rewritingManager);
                refactoring.Refactor(selection.Value);
            }
        }
    }
}
