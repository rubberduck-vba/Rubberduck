using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.IdentifierReferences
{
    public class ImplementationsListDockablePresenter : DockableToolwindowPresenter
    {
        public ImplementationsListDockablePresenter(IVBE vbe, IAddIn addin, IDockableUserControl control, IEnumerable<Declaration> implementations)
            : base(vbe, addin, control, null)
        {
            BindTarget(implementations);
        }

        private void BindTarget(IEnumerable<Declaration> implementations)
        {
            var control = UserControl as SimpleListControl;
            Debug.Assert(control != null);

            var listBox = control.ResultBox;
            listBox.DataSource = implementations.Select(implementation => new ImplementationListItem(implementation)).ToList();
            listBox.DisplayMember = "DisplayString";
            listBox.ValueMember = "Selection";
            control.Navigate += ControlNavigate;
        }

        public static void OnNavigateImplementation(Declaration implementation)
        {
            implementation.QualifiedName.QualifiedModuleName.Component.CodeModule.CodePane.Selection = implementation.Selection;
        }

        private void ControlNavigate(object sender, ListItemActionEventArgs e)
        {
            var implementation = e.SelectedItem as ImplementationListItem;
            if (implementation != null)
            {
                OnNavigateImplementation(implementation.GetDeclaration());
            }
        }
    }
}
