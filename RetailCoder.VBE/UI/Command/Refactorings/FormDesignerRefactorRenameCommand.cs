using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Rename;
using Rubberduck.UI.Refactorings;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class FormDesignerRefactorRenameCommand : RefactorCommandBase
    {
        private readonly VBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly ICodePaneWrapperFactory _wrapperWrapperFactory;

        public FormDesignerRefactorRenameCommand(VBE vbe, RubberduckParserState state, ICodePaneWrapperFactory wrapperWrapperFactory) 
            : base (vbe)
        {
            _vbe = vbe;
            _state = state;
            _wrapperWrapperFactory = wrapperWrapperFactory;
        }

        public override bool CanExecute(object parameter)
        {
            return _state.Status == ParserState.Ready;
        }

        public override void Execute(object parameter)
        {
            using (var view = new RenameDialog())
            {
                var factory = new RenamePresenterFactory(Vbe, view, _state, new MessageBox(), _wrapperWrapperFactory);
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