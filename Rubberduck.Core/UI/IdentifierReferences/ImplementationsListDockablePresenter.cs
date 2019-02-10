using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.IdentifierReferences
{
    public class ImplementationsListDockablePresenter : DockableToolwindowPresenter
    {
        private readonly ISelectionService _selectionService;

        public ImplementationsListDockablePresenter(IVBE vbe, IAddIn addin, IDockableUserControl control, ISelectionService selectionService, IEnumerable<Declaration> implementations)
            : base(vbe, addin, control, null)
        {
            _selectionService = selectionService;

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

        private void OnNavigateImplementation(Declaration implementation)
        {
            _selectionService.TrySetActiveSelection(implementation.QualifiedModuleName, implementation.Selection);
        }

        private void ControlNavigate(object sender, ListItemActionEventArgs e)
        {
            if (e.SelectedItem is ImplementationListItem implementation)
            {
                OnNavigateImplementation(implementation.GetDeclaration());
            }
        }
    }
}
