using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ReorderParameters;
using Rubberduck.UI.Refactorings.ReorderParameters;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorReorderParametersCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _msgbox;

        public RefactorReorderParametersCommand(IVBE vbe, RubberduckParserState state, IMessageBox msgbox) 
            : base (vbe)
        {
            _state = state;
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

        protected override bool CanExecuteImpl(object parameter)
        {
            var pane = Vbe.ActiveCodePane;
            {
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
                var canExecute = (member.DeclarationType == DeclarationType.PropertyLet || member.DeclarationType == DeclarationType.PropertySet)
                        ? parameters.Count > 2
                        : parameters.Count > 1;

                return canExecute;
            }
        }

        protected override void ExecuteImpl(object parameter)
        {
            var pane = Vbe.ActiveCodePane;
            var module = pane.CodeModule;
            {
                if (pane.IsWrappingNullReference)
                {
                    return;
                }
                var selection = new QualifiedSelection(new QualifiedModuleName(module.Parent), pane.Selection);

                using (var view = new ReorderParametersDialog(new ReorderParametersViewModel(_state)))
                {
                    var factory = new ReorderParametersPresenterFactory(Vbe, view, _state, _msgbox);
                    var refactoring = new ReorderParametersRefactoring(Vbe, factory, _msgbox);
                    refactoring.Refactor(selection);
                }
            }
        }
    }
}
