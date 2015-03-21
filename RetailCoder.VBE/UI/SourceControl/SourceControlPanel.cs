using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Rubberduck.UI.SourceControl
{
    [SuppressMessage("ReSharper", "ArrangeThisQualifier")]
    public partial class SourceControlPanel : UserControl, ISourceControlView
    {
        public SourceControlPanel()
        {
            InitializeComponent();
        }

        public string ClassId
        {
            get { return "19A32FC9-4902-4385-9FE7-829D4F9C441D"; }
        }

        public string Caption
        {
            get { return "Source Control"; }
        }


        public string CommitMessage
        {
            get { return this.CommitMessageBox.Text; }
            set { this.CommitMessageBox.Text = value; }
        }

        public CommitAction CommitAction
        {
            get { return (CommitAction)this.CommitActionDropdown.SelectedIndex; }
            set { this.CommitActionDropdown.SelectedIndex = (int)value; }
        }


        private BindingList<string> _includedChanges = new BindingList<string>();
        public IList<string> IncludedChanges
        {
            get { return _includedChanges; }
            set { _includedChanges = new BindingList<string>(value); }
        }

        private BindingList<string> _excludedChanges = new BindingList<string>();
        public IList<string> ExcludedChanges
        {
            get { return _excludedChanges; }
            set { _excludedChanges = new BindingList<string>(value); }
        }

        private BindingList<string> _untrackedFiles = new BindingList<string>(); 
        public IList<string> UntrackedFiles
        {
            get { return _untrackedFiles; }
            set { _untrackedFiles = new BindingList<string>(value);}
        }

        public event System.EventHandler<System.EventArgs> Commit;
        public event System.EventHandler<System.EventArgs> RefreshData;
    }
}
