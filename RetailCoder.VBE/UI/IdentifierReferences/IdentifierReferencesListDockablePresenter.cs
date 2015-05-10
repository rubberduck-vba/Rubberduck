using System;
using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;

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
            Control.NavigateIdentifierReference += Control_NavigateIdentifierReference;
        }

        public static void OnNavigateIdentifierReference(VBE vbe, IdentifierReference reference)
        {
            vbe.SetSelection(new QualifiedSelection(reference.QualifiedModuleName, reference.Selection));
        }

        private void Control_NavigateIdentifierReference(object sender, NavigateCodeEventArgs e)
        {
            OnNavigateIdentifierReference(VBE, e.Reference);
        }

        IdentifierReferencesListControl Control { get { return UserControl as IdentifierReferencesListControl; } }
    }
}
