using System.Diagnostics;
using System.Linq;
using Microsoft.Vbe.Interop;
using System.Runtime.InteropServices;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ImplementInterface;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

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
            return true;
            if (Vbe.ActiveCodePane == null || _state.Status != ParserState.Ready)
            {
                return false;
            }

            //var target = _state.FindSelectedDeclaration(Vbe.ActiveCodePane); // nope. logic is a bit more complex here.

            var selection = Vbe.ActiveCodePane.GetSelection();
            var targetInterface = _state.AllUserDeclarations.FindInterface(selection);

            var targetClass = _state.AllUserDeclarations.SingleOrDefault(d =>
                        !d.IsBuiltIn && d.DeclarationType == DeclarationType.ClassModule &&
                        d.QualifiedSelection.QualifiedName.Equals(selection.QualifiedName));

            var canExecute = targetInterface != null && targetClass != null;

            Debug.WriteLine("{0}.CanExecute evaluates to {1}", GetType().Name, canExecute);
            return canExecute;
        }

        public override void Execute(object parameter)
        {
            if (Vbe.ActiveCodePane == null)
            {
                return;
            }

            var refactoring = new ImplementInterfaceRefactoring(_state, Editor, new MessageBox());
            refactoring.Refactor();
        }
    }
}