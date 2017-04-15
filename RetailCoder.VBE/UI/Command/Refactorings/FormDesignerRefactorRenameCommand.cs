using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Rename;
using Rubberduck.UI.Refactorings.Rename;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class FormDesignerRefactorRenameCommand : RefactorCommandBase
    {
        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _messageBox;

        public FormDesignerRefactorRenameCommand(IVBE vbe, RubberduckParserState state, IMessageBox messageBox) 
            : base (vbe)
        {
            _vbe = vbe;
            _state = state;
            _messageBox = messageBox;
        }

        protected override bool CanExecuteImpl(object parameter)
        {
            return _state.Status == ParserState.Ready;
        }

        protected override void ExecuteImpl(object parameter)
        {
            using (var view = new RenameDialog(new RenameViewModel(_state)))
            {
                var factory = new RenamePresenterFactory(Vbe, view, _state);
                var refactoring = new RenameRefactoring(Vbe, factory, _messageBox, _state);

                var target = GetTarget();

                if (target != null)
                {
                    refactoring.Refactor(target);
                }
            }
        }

        private Declaration GetTarget()
        {
            var project = _vbe.ActiveVBProject;
            var component = _vbe.SelectedVBComponent;
            {
                if (Vbe.SelectedVBComponent != null && Vbe.SelectedVBComponent.HasDesigner)
                {
                    var designer = ((dynamic)component.Target).Designer;

                    if (designer.selected.count == 1)
                    {
                        var control = designer.selected.item(0);
                        var result = _state.AllUserDeclarations
                            .FirstOrDefault(item => item.DeclarationType == DeclarationType.Control
                                                    && project.HelpFile == item.ProjectId
                                                    && item.ComponentName == component.Name
                                                    && item.IdentifierName == control.Name);

                        Marshal.ReleaseComObject(control);
                        Marshal.ReleaseComObject(designer);
                        return result;
                    } else {
                        var message = string.Format(RubberduckUI.RenameDialog_AmbiguousSelection);
                        _messageBox.Show(message, RubberduckUI.RenameDialog_Caption, MessageBoxButtons.OK,
                            MessageBoxIcon.Exclamation);
                    }
                }
            }

            return null;
        }
    }
}
