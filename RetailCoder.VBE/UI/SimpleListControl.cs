using System;
using System.Windows.Forms;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI.IdentifierReferences;

namespace Rubberduck.UI
{
    public partial class SimpleListControl : UserControl, IDockableUserControl
    {
        public SimpleListControl(Declaration target)
            : this(string.Format(RubberduckUI.AllReferences_Caption, target.IdentifierName))
        { }

        public SimpleListControl(string caption)
        {
            _caption = caption;
            InitializeComponent();
            ResultBox.DoubleClick += ResultBox_DoubleClick;
        }

        public event EventHandler<ListItemActionEventArgs> Navigate;
        private void ResultBox_DoubleClick(object sender, EventArgs e)
        {
            var handler = Navigate;
            if (handler == null || ResultBox.SelectedItem == null)
            {
                return;
            }

            var arg = new ListItemActionEventArgs(ResultBox.SelectedItem);
            handler(this, arg);
        }

        private const string ClassId = "972A7CE8-55A0-48F5-B607-2035E81D28CF";
        string IDockableUserControl.ClassId { get { return ClassId; } }

        private readonly string _caption;
        string IDockableUserControl.Caption
        {
            get { return _caption; }
        }

        public ViewModelBase ViewModel { get; set; }
    }
}
