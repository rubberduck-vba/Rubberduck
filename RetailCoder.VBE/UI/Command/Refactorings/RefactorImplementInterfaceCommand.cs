using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ImplementInterface;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorImplementInterfaceCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _msgBox;

        public RefactorImplementInterfaceCommand(IVBE vbe, RubberduckParserState state, IMessageBox msgBox)
            : base(vbe)
        {
            _state = state;
            _msgBox = msgBox;
        }

        protected override bool CanExecuteImpl(object parameter)
        {
            var pane = Vbe.ActiveCodePane;
            {
                if (_state.Status != ParserState.Ready || pane.IsWrappingNullReference)
                {
                    return false;
                }

                var selection = pane.GetQualifiedSelection();
                if (!selection.HasValue)
                {
                    return false;
                }

                var targetInterface = _state.AllUserDeclarations.FindInterface(selection.Value);

                var targetClass = _state.AllUserDeclarations.SingleOrDefault(d =>
                    d.DeclarationType == DeclarationType.ClassModule &&
                    d.QualifiedSelection.QualifiedName.Equals(selection.Value.QualifiedName));

                return targetInterface != null && targetClass != null;
            }
        }

        protected override void ExecuteImpl(object parameter)
        {
            var pane = Vbe.ActiveCodePane;
            {
                if (pane.IsWrappingNullReference)
                {
                    return;
                }

                var refactoring = new ImplementInterfaceRefactoring(Vbe, _state, _msgBox);
                refactoring.Refactor();
            }
        }
    }
}
