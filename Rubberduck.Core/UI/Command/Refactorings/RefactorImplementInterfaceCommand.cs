using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ImplementInterface;
using Rubberduck.VBEditor;
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

        protected override bool EvaluateCanExecute(object parameter)
        {
            
            var selection = Vbe.GetActiveSelection();        

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

        protected override void OnExecute(object parameter)
        {
            using (var pane = Vbe.ActiveCodePane)
            {
                if (pane.IsWrappingNullReference)
                {
                    return;
                }
            }
            var refactoring = new ImplementInterfaceRefactoring(Vbe, _state, _msgBox);
            refactoring.Refactor();
        }
    }
}
