using System.Collections.Generic;
using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.UI.IdentifierReferences
{
    public class ImplementationsListDockablePresenter : DockableToolwindowPresenter
    {
        private static ICodePaneWrapperFactory _wrapperFactory;

        public ImplementationsListDockablePresenter(VBE vbe, AddIn addin, SimpleListControl control, IEnumerable<Declaration> implementations, ICodePaneWrapperFactory wrapperFactory)
            : base(vbe, addin, control)
        {
            BindTarget(implementations);
            _wrapperFactory = wrapperFactory;
        }

        private void BindTarget(IEnumerable<Declaration> implementations)
        {
            var listBox = Control.ResultBox;
            listBox.DataSource = implementations.Select(implementation => new ImplementationListItem(implementation, _wrapperFactory)).ToList();
            listBox.DisplayMember = "DisplayString";
            listBox.ValueMember = "Selection";
            Control.Navigate += ControlNavigate;
        }

        public static void OnNavigateImplementation(VBE vbe, Declaration implementation)
        {
            vbe.SetSelection(implementation.QualifiedName.QualifiedModuleName.Project, implementation.Selection, implementation.QualifiedName.QualifiedModuleName.Component.Name, _wrapperFactory);
        }

        private void ControlNavigate(object sender, ListItemActionEventArgs e)
        {
            var implementation = e.SelectedItem as ImplementationListItem;
            if (implementation != null)
            {
                OnNavigateImplementation(VBE, implementation.GetDeclaration());
            }
        }

        SimpleListControl Control { get { return UserControl as SimpleListControl; } }
    }
}