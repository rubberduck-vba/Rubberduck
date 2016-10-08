using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.SafeComWrappers.VBA.Abstract;

namespace Rubberduck.UI.IdentifierReferences
{
    public class ImplementationsListDockablePresenter : DockableToolwindowPresenter
    {
        public ImplementationsListDockablePresenter(IVBE vbe, IAddIn addin, SimpleListControl control, IEnumerable<Declaration> implementations)
            : base(vbe, addin, control)
        {
            BindTarget(implementations);
        }

        private void BindTarget(IEnumerable<Declaration> implementations)
        {
            var listBox = Control.ResultBox;
            listBox.DataSource = implementations.Select(implementation => new ImplementationListItem(implementation)).ToList();
            listBox.DisplayMember = "DisplayString";
            listBox.ValueMember = "Selection";
            Control.Navigate += ControlNavigate;
        }

        public static void OnNavigateImplementation(Declaration implementation)
        {
            implementation.QualifiedName.QualifiedModuleName.Component.CodeModule.CodePane.SetSelection(implementation.Selection);
        }

        private void ControlNavigate(object sender, ListItemActionEventArgs e)
        {
            var implementation = e.SelectedItem as ImplementationListItem;
            if (implementation != null)
            {
                OnNavigateImplementation(implementation.GetDeclaration());
            }
        }

        SimpleListControl Control { get { return UserControl as SimpleListControl; } }
    }
}
