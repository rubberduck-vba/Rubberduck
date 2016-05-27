using System.Linq;
using Microsoft.Vbe.Interop;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Rename;
using Rubberduck.UI.Refactorings;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class FormDesignerRefactorRenameCommand : RefactorCommandBase
    {
        private readonly VBE _vbe;
        private readonly RubberduckParserState _state;

        public FormDesignerRefactorRenameCommand(VBE vbe, RubberduckParserState state) 
            : base (vbe)
        {
            _vbe = vbe;
            _state = state;
        }

        public override bool CanExecute(object parameter)
        {
            return _state.Status == ParserState.Ready;
        }

        public override void Execute(object parameter)
        {
            using (var view = new RenameDialog())
            {
                var factory = new RenamePresenterFactory(Vbe, view, _state, new MessageBox());
                var refactoring = new RenameRefactoring(Vbe, factory, new MessageBox(), _state);

                var target = GetTarget();

                if (target != null)
                {
                    refactoring.Refactor(target);
                }
            }
        }

        private Declaration GetTarget()
        {
            if (Vbe.SelectedVBComponent != null && Vbe.SelectedVBComponent.Designer != null)
            {
                var designer = (dynamic)Vbe.SelectedVBComponent.Designer;

                foreach (var control in designer.Controls)
                {
                    if (!control.InSelection)
                    {
                        continue;
                    }

                    return _state.AllUserDeclarations
                        .FirstOrDefault(item => item.DeclarationType == DeclarationType.Control
                            && Vbe.ActiveVBProject.HelpFile == item.ProjectId
                            && item.ComponentName == Vbe.SelectedVBComponent.Name
                            && item.IdentifierName == control.Name);
                }
            }

            return null;
        }
    }
}
