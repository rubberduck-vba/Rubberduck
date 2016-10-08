using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.SafeComWrappers.VBA.Abstract;

namespace Rubberduck.UI.IdentifierReferences
{
    public class IdentifierReferencesListDockablePresenter : DockableToolwindowPresenter
    {
        public IdentifierReferencesListDockablePresenter(IVBE vbe, IAddIn addin, SimpleListControl control, Declaration target)
            : base(vbe, addin, control)
        {
            BindTarget(target);
        }

        private void BindTarget(Declaration target)
        {
            var listBox = Control.ResultBox;
            listBox.DataSource = target.References.Select(reference => new IdentifierReferenceListItem(reference)).ToList();
            listBox.DisplayMember = "DisplayString";
            listBox.ValueMember = "Selection";
            Control.Navigate += ControlNavigate;
        }

        public static void OnNavigateIdentifierReference(IdentifierReference reference)
        {
            reference.QualifiedModuleName.Component.CodeModule.CodePane.SetSelection(reference.Selection);
        }

        private void ControlNavigate(object sender, ListItemActionEventArgs e)
        {
            var reference = e.SelectedItem as IdentifierReferenceListItem;
            if (reference != null)
            {
                OnNavigateIdentifierReference(reference.GetReferenceItem());
            }
        }

        SimpleListControl Control { get { return UserControl as SimpleListControl; } }
    }
}
