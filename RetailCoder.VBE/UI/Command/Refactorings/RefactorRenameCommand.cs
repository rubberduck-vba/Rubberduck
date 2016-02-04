using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Rename;
using Rubberduck.UI.Refactorings;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorRenameCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly ICodePaneWrapperFactory _wrapperWrapperFactory;

        public RefactorRenameCommand(VBE vbe, RubberduckParserState state, IActiveCodePaneEditor editor, ICodePaneWrapperFactory wrapperWrapperFactory) 
            : base (vbe, editor)
        {
            _state = state;
            _wrapperWrapperFactory = wrapperWrapperFactory;
        }

        public override void Execute(object parameter)
        {
            using (var view = new RenameDialog())
            {
                var factory = new RenamePresenterFactory(Vbe, view, _state, new MessageBox(), _wrapperWrapperFactory);
                var refactoring = new RenameRefactoring(factory, Editor, new MessageBox(), _state);

                var target = GetTarget(parameter);

                if (target == null)
                {
                    refactoring.Refactor();
                }
                else
                {
                    refactoring.Refactor(target);
                }
            }
        }

        private Declaration GetTarget(object parameter)
        {
            var target = parameter as Declaration;
            if (target != null)
            {
                return target;
            }

            // rename project
            if (Vbe.SelectedVBComponent == null)
            {
                return
                    _state.AllUserDeclarations.SingleOrDefault(d =>
                            d.DeclarationType == DeclarationType.Project && d.IdentifierName == Vbe.ActiveVBProject.Name);
            }

            // selected component is not active
            if (Vbe.ActiveCodePane == null || Vbe.ActiveCodePane.CodeModule != Vbe.SelectedVBComponent.CodeModule)
            {
                // selected pane is userform - see if there are selected controls
                if (Vbe.SelectedVBComponent.Designer != null)
                {
                    var designer = (dynamic)Vbe.SelectedVBComponent.Designer;

                    foreach (var control in designer.Controls)
                    {
                        if (!control.InSelection)
                        {
                            continue;
                        }

                        target = _state.AllUserDeclarations
                            .FirstOrDefault(item => item.IdentifierName == control.Name &&
                                                    item.ComponentName == Vbe.SelectedVBComponent.Name &&
                                                    Vbe.ActiveVBProject.Equals(item.Project));

                        break;
                    }
                }

                // user form is not designer or there were no selected controls
                if (target == null)
                {
                    target = _state.AllUserDeclarations.SingleOrDefault(
                        t => t.IdentifierName == Vbe.SelectedVBComponent.Name &&
                             t.Project == Vbe.ActiveVBProject &&
                             new[]
                                 {
                                     DeclarationType.Class,
                                     DeclarationType.Document,
                                     DeclarationType.Module,
                                     DeclarationType.UserForm
                                 }.Contains(t.DeclarationType));
                }
            }

            return target;
        }
    }
}