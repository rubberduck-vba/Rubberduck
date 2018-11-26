using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.IdentifierReferences
{
    public class ImplementationsListDockablePresenter : DockableToolwindowPresenter
    {
        private readonly IProjectsProvider _projectsProvider;

        public ImplementationsListDockablePresenter(IVBE vbe, IAddIn addin, IDockableUserControl control, IProjectsProvider projectsProvider, IEnumerable<Declaration> implementations)
            : base(vbe, addin, control, null)
        {
            _projectsProvider = projectsProvider;

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
            var component = _projectsProvider.Component(implementation.QualifiedName.QualifiedModuleName);
            using (var codeModule = component.CodeModule)
            {
                using (var codePane = codeModule.CodePane)
                {
                    codePane.Selection = implementation.Selection;
                }
            }
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
