using System.Windows.Forms;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.UI.IdentifierReferences
{
    public partial class IdentifierReferencesListControl : UserControl, IDockableUserControl
    {
        public IdentifierReferencesListControl(Declaration target)
        {
            InitializeComponent();
            Target = target;
        }

        public Declaration Target { get; private set; }

        private const string ClassId = "972A7CE8-55A0-48F5-B607-2035E81D28CF";
        string IDockableUserControl.ClassId { get { return ClassId; } }
        string IDockableUserControl.Caption { get { return string.Format(RubberduckUI.AllReferences_Caption, Target.IdentifierName); } }
    }
}
