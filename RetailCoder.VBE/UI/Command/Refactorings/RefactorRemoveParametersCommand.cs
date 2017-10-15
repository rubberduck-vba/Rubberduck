using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.RemoveParameters;
using Rubberduck.UI.Refactorings.RemoveParameters;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorRemoveParametersCommand : RefactorCommandBase
    {
        private readonly IMessageBox _msgbox;
        private readonly RubberduckParserState _state;

        public RefactorRemoveParametersCommand(IVBE vbe, RubberduckParserState state, IMessageBox msgbox) 
            : base (vbe)
        {
            _msgbox = msgbox;
            _state = state;
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

        protected override bool CanExecuteImpl(object parameter)
        {
            var pane = Vbe.ActiveCodePane;
            if (pane.IsWrappingNullReference || _state.Status != ParserState.Ready)
            {
                return false;
            }

            var selection = pane.GetQualifiedSelection();
            var member = _state.AllUserDeclarations.FindTarget(selection.Value, ValidDeclarationTypes);
            if (member == null)
            {
                return false;
            }

            var parameters = _state.AllUserDeclarations.Where(item => item.DeclarationType == DeclarationType.Parameter && member.Equals(item.ParentScopeDeclaration)).ToList();
            return member.DeclarationType == DeclarationType.PropertyLet || member.DeclarationType == DeclarationType.PropertySet
                    ? parameters.Count > 1
                    : parameters.Any();
        }

        protected override void ExecuteImpl(object parameter)
        {
            var pane = Vbe.ActiveCodePane;
            if (pane.IsWrappingNullReference)
            {
                return;
            }

            var selection = pane.GetQualifiedSelection();
            using (var view = new RemoveParametersDialog(new RemoveParametersViewModel(_state)))
            {
                var factory = new RemoveParametersPresenterFactory(Vbe, view, _state, _msgbox);
                var refactoring = new RemoveParametersRefactoring(Vbe, factory);
                refactoring.Refactor(selection.Value);
            }
        }
    }
}
