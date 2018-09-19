using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Interaction;
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

        protected override bool EvaluateCanExecute(object parameter)
        {
            
            var selection = Vbe.GetActiveSelection();        

            if (!selection.HasValue)
            {
                return false;
            }

            var targetInterface = _state.DeclarationFinder.FindInterface(selection.Value);
            
            var targetClass = _state.DeclarationFinder.Members(selection.Value.QualifiedName)
                .SingleOrDefault(declaration => declaration.DeclarationType == DeclarationType.ClassModule);

            return targetInterface != null && targetClass != null
                && !_state.IsNewOrModified(targetInterface.QualifiedModuleName)
                && !_state.IsNewOrModified(targetClass.QualifiedModuleName);
            
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
