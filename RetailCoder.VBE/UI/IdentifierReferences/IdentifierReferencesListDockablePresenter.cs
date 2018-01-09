using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.IdentifierReferences
{
    public class IdentifierReferencesListDockablePresenter : DockableToolwindowPresenter
    {
        public IdentifierReferencesListDockablePresenter(IVBE vbe, IAddIn addin, SimpleListControl control, Declaration target)
            : base(vbe, addin, control, null)
        {
            BindTarget(target);
        }

        private void BindTarget(Declaration target)
        {
            var control = UserControl as SimpleListControl;
            if (control == null) { return; }

            var listBox = control.ResultBox;
            listBox.DataSource = target.References.Select(reference => new IdentifierReferenceListItem(reference)).ToList();
            listBox.DisplayMember = "DisplayString";
            listBox.ValueMember = "Selection";
            control.Navigate += ControlNavigate;
        }

        public static void OnNavigateIdentifierReference(IdentifierReference reference)
        {
            using (var codeModule = reference.QualifiedModuleName.Component.CodeModule)
            {
                using (var codePane = codeModule.CodePane)
                {
                    codePane.Selection = reference.Selection;
                }
            }
        }

        private void ControlNavigate(object sender, ListItemActionEventArgs e)
        {
            if (e.SelectedItem is IdentifierReferenceListItem reference)
            {
                OnNavigateIdentifierReference(reference.GetReferenceItem());
            }
        }
    }
}
