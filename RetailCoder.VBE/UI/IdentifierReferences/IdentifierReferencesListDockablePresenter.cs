using System.Collections;
using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.UI.IdentifierReferences
{
    public class IdentifierReferencesListDockablePresenter : DockablePresenterBase
    {
        public IdentifierReferencesListDockablePresenter(VBE vbe, AddIn addin, SimpleListControl control, Declaration target)
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

        public static void OnNavigateIdentifierReference(VBE vbe, IdentifierReference reference)
        {
            vbe.SetSelection(new QualifiedSelection(reference.QualifiedModuleName, reference.Selection));
        }

        private void ControlNavigate(object sender, ListItemActionEventArgs e)
        {
            var reference = e.SelectedItem as IdentifierReferenceListItem;
            if (reference != null)
            {
                OnNavigateIdentifierReference(VBE, reference.GetReferenceItem());
            }
        }

        SimpleListControl Control { get { return UserControl as SimpleListControl; } }
    }
}
