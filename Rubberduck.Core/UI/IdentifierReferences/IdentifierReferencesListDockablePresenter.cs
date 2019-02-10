using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.IdentifierReferences
{
    public class IdentifierReferencesListDockablePresenter : DockableToolwindowPresenter
    {
        private readonly ISelectionService _selectionService;

        public IdentifierReferencesListDockablePresenter(IVBE vbe, IAddIn addin, SimpleListControl control, ISelectionService selectionService, Declaration target)
            : base(vbe, addin, control, null)
        {
            _selectionService = selectionService;

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

        private void OnNavigateIdentifierReference(IdentifierReference reference)
        {
            _selectionService.TrySetActiveSelection(reference.QualifiedModuleName, reference.Selection);
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
