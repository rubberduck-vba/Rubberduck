using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.UI.IdentifierReferences
{
    public class IdentifierReferenceListItem
    {
        private readonly IdentifierReference _reference;

        public IdentifierReferenceListItem(IdentifierReference reference)
        {
            _reference = reference;
        }

        public QualifiedSelection Selection { get { return new QualifiedSelection(_reference.QualifiedModuleName, _reference.Selection); } }
        public string IdentifierName { get { return _reference.IdentifierName; } }
    }

    public class IdentifierReferencesListDockablePresenter : DockablePresenterBase
    {
        public IdentifierReferencesListDockablePresenter(VBE vbe, AddIn addin, IdentifierReferencesListControl control, Declaration target) 
            : base(vbe, addin, control)
        {
            var listBox = ((IdentifierReferencesListControl) UserControl).ResultBox;

            listBox.DataSource = target.References.Select(reference => new IdentifierReferenceListItem(reference));
            listBox.DisplayMember = "IdentifierName";
            listBox.ValueMember = "Selection";
        }
    }
}
