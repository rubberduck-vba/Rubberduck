using System;
using System.Windows.Forms;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.UI.IdentifierReferences
{
    public partial class IdentifierReferencesListControl : UserControl, IDockableUserControl
    {
        private readonly string _caption;

        public IdentifierReferencesListControl(Declaration target)
            : this(target, string.Format(RubberduckUI.AllReferences_Caption, target.IdentifierName))
        { }

        public IdentifierReferencesListControl(Declaration target, string caption)
        {
            _caption = caption;
            InitializeComponent();
            Target = target;
            ResultBox.DoubleClick += ResultBox_DoubleClick;
        }

        public event EventHandler<NavigateCodeEventArgs> NavigateIdentifierReference;
        private void ResultBox_DoubleClick(object sender, EventArgs e)
        {
            var handler = NavigateIdentifierReference;
            if (handler == null || ResultBox.SelectedItem == null)
            {
                return;
            }

            var selectedItem = ResultBox.SelectedItem as IdentifierReferenceListItem;
            if (selectedItem == null)
            {
                return;
            }

            var arg = new NavigateCodeEventArgs(selectedItem.GetReferenceItem());
            handler(this, arg);
        }

        public Declaration Target { get; private set; }

        private const string ClassId = "972A7CE8-55A0-48F5-B607-2035E81D28CF";
        string IDockableUserControl.ClassId { get { return ClassId; } }

        string IDockableUserControl.Caption
        {
            get { return _caption; }
        }
    }
}
