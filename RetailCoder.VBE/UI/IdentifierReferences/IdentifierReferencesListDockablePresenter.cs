using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.UI.IdentifierReferences
{
    public class IdentifierReferencesListDockablePresenter : DockablePresenterBase
    {
        private static IRubberduckCodePaneFactory _factory;

        public IdentifierReferencesListDockablePresenter(VBE vbe, AddIn addin, SimpleListControl control, Declaration target, IRubberduckCodePaneFactory factory)
            : base(vbe, addin, control)
        {
            _factory = factory;
            BindTarget(target);
        }

        private void BindTarget(Declaration target)
        {
            var listBox = Control.ResultBox;
            listBox.DataSource = target.References.Select(reference => new IdentifierReferenceListItem(reference, _factory)).ToList();
            listBox.DisplayMember = "DisplayString";
            listBox.ValueMember = "Selection";
            Control.Navigate += ControlNavigate;
        }

        public static void OnNavigateIdentifierReference(VBE vbe, IdentifierReference reference)
        {
            vbe.SetSelection(reference.QualifiedModuleName.Project, reference.Selection, reference.QualifiedModuleName.Component.Name, _factory);
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
