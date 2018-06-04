using System;
using System.Windows.Forms;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI.IdentifierReferences;
using Rubberduck.Resources;

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

        private readonly string RandomGuid = Guid.NewGuid().ToString();
        string IDockableUserControl.GuidIdentifier => RandomGuid;

        private readonly string _caption;
        string IDockableUserControl.Caption
        {
            get { return _caption; }
        }

        public ViewModelBase ViewModel { get; set; }
    }
}
