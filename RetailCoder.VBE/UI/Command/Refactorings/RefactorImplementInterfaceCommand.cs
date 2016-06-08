using System.Linq;
using Microsoft.Vbe.Interop;
using System.Runtime.InteropServices;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ImplementInterface;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorImplementInterfaceCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;

        public RefactorImplementInterfaceCommand(VBE vbe, RubberduckParserState state)
            : base(vbe)
        {
            _state = state;
        }

        public override bool CanExecute(object parameter)
        {
            if (Vbe.ActiveCodePane == null || _state.Status != ParserState.Ready)
            {
                return false;
            }

            var selection = Vbe.ActiveCodePane.GetQualifiedSelection();
            if (!selection.HasValue)
            {
                return false;
            }

            var targetInterface = _state.AllUserDeclarations.FindInterface(selection.Value);

            var targetClass = _state.AllUserDeclarations.SingleOrDefault(d =>
                        !d.IsBuiltIn && d.DeclarationType == DeclarationType.ClassModule &&
                        d.QualifiedSelection.QualifiedName.Equals(selection.Value.QualifiedName));

            var canExecute = targetInterface != null && targetClass != null;

            return canExecute;
        }

        public override void Execute(object parameter)
        {
            if (Vbe.ActiveCodePane == null)
            {
                return;
            }

            var refactoring = new ImplementInterfaceRefactoring(Vbe, _state, new MessageBox());
            refactoring.Refactor();
        }
    }
}
