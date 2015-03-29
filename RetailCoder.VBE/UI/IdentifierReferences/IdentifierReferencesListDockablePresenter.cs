using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
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
        private ListBox ReferencesList { get { return ((IdentifierReferencesListControl) UserControl).ResultBox; } }

        public IdentifierReferencesListDockablePresenter(VBE vbe, AddIn addin, IdentifierReferencesListControl control, Declaration target) 
            : base(vbe, addin, control)
        {
            // todo: change to gridview - this listbox is just to see something work.
            ReferencesList.DataSource = target.References.Select(reference => new IdentifierReferenceListItem(reference)).ToList();
            ReferencesList.DisplayMember = "Selection";
            ReferencesList.ValueMember = "Selection";
            ReferencesList.Refresh();

            ReferencesList.DoubleClick += ReferencesList_DoubleClick;
        }

        private void ReferencesList_DoubleClick(object sender, EventArgs e)
        {
            if (ReferencesList.SelectedItem == null)
            {
                return;
            }

            VBE.SetSelection((QualifiedSelection)ReferencesList.SelectedItem);
        }
    }
}
