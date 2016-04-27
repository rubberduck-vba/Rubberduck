using System.Linq;
using Microsoft.Vbe.Interop;
using System.Runtime.InteropServices;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ImplementInterface;
using Rubberduck.VBEditor;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorImplementInterfaceCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;

        public RefactorImplementInterfaceCommand(VBE vbe, RubberduckParserState state, IActiveCodePaneEditor editor)
            : base(vbe, editor)
        {
            _state = state;
        }

        public override bool CanExecute(object parameter)
        {
            if (Vbe.ActiveCodePane == null || _state.Status != ParserState.Ready)
            {
                return false;
            }

            var selection = Editor.GetSelection();

            if (!selection.HasValue)
            {
                return false;
            }

            var targetInterface = _state.AllUserDeclarations.FindInterface(selection.Value);

            var targetClass = _state.AllUserDeclarations.SingleOrDefault(d =>
                        !d.IsBuiltIn && d.DeclarationType == DeclarationType.ClassModule &&
                        d.QualifiedSelection.QualifiedName.Equals(selection.Value.QualifiedName));

            return targetClass != null && targetInterface != null;
        }

        public override void Execute(object parameter)
        {
            var refactoring = new ImplementInterfaceRefactoring(_state, Editor, new MessageBox());

            // ReSharper disable once PossibleInvalidOperationException
            refactoring.Refactor(Editor.GetSelection().Value);
        }
    }
}