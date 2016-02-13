using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.UI.IdentifierReferences
{
    public class IdentifierReferencesListDockablePresenter : DockableToolwindowPresenter
    {
        private static ICodePaneWrapperFactory _wrapperFactory;

        public IdentifierReferencesListDockablePresenter(VBE vbe, AddIn addin, SimpleListControl control, Declaration target, ICodePaneWrapperFactory wrapperFactory)
            : base(vbe, addin, control)
        {
            _wrapperFactory = wrapperFactory;
            BindTarget(target);
        }

        private void BindTarget(Declaration target)
        {
            var listBox = Control.ResultBox;
            listBox.DataSource = target.References.Select(reference => new IdentifierReferenceListItem(reference, _wrapperFactory)).ToList();
            listBox.DisplayMember = "DisplayString";
            listBox.ValueMember = "Selection";
            Control.Navigate += ControlNavigate;
        }

        public static void OnNavigateIdentifierReference(VBE vbe, IdentifierReference reference)
        {
            vbe.SetSelection(reference.QualifiedModuleName.Project, reference.Selection, reference.QualifiedModuleName.Component.Name, _wrapperFactory);
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
