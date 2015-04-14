using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.UI.IdentifierReferences
{
    public class IdentifierReferencesListDockablePresenter : DockablePresenterBase
    {
        public IdentifierReferencesListDockablePresenter(VBE vbe, AddIn addin, IdentifierReferencesListControl control, Declaration target)
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
        }

        IdentifierReferencesListControl Control { get { return UserControl as IdentifierReferencesListControl; } }
    }
}
