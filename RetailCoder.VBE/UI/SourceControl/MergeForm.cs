using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Rubberduck.UI.SourceControl
{
    public partial class MergeForm : Form, IMergeView
    {
        public MergeForm()
        {
            InitializeComponent();
        }

        public bool OkayButtonEnabled
        {
            get { return this.OkayButton.Enabled; }
            set { this.OkayButton.Enabled = value; }
        }

        private BindingList<string> _source;
        public IList<string> SourceSelectorData
        {
            get { return _source; }
            set
            {
                _source = new BindingList<string>(value);
                this.SourceSelector.DataSource = _source;
            }
        }

        private BindingList<string> _destination;
        public IList<string> DestinationSelectorData
        {
            get { return _destination; }
            set
            {
                _destination = new BindingList<string>(value);
                this.DestinationSelector.DataSource = _destination;
            }
        }

        public string SelectedSourceBranch
        {
            get { return this.SourceSelector.SelectedItem.ToString(); }
            set { this.SourceSelector.SelectedItem = value; }
        }

        public string SelectedDestinationBranch
        {
            get { return this.DestinationSelector.SelectedItem.ToString(); }
            set { this.DestinationSelector.SelectedItem = value; }
        }

        public string StatusText
        {
            get { return this.StatusTextBox.Text; }
            set { this.StatusTextBox.Text = value; }
        }

        public bool StatusTextVisible
        {
            get { return this.StatusTextBox.Visible; }
            set { this.StatusTextBox.Visible = value; }
        }

        public event EventHandler<EventArgs> MergeStatusChanged;

        private MergeStatus _mergeStatus;
        public MergeStatus Status
        {
            get { return _mergeStatus; }
            set
            {
                _mergeStatus = value;

                var handler = MergeStatusChanged;
                if (handler != null)
                {
                    handler(this, EventArgs.Empty);
                }
            }
        }

        public event EventHandler<EventArgs> Confirm;
        private void OnConfirm(object sender, EventArgs e)
        {
            var handler = Confirm;
            if (handler != null)
            {
                handler(this, e);
            }
        }

        public event EventHandler<EventArgs> Cancel;
        private void OnCancel(object sender, EventArgs e)
        {
            var handler = Cancel;
            if (handler != null)
            {
                handler(this, e);
            }
        }

        public event EventHandler<EventArgs> SelectedSourceBranchChanged;

        private void OnSelectedSourceBranchChanged(object sender, EventArgs e)
        {
            var handler = SelectedSourceBranchChanged;
            if (handler != null)
            {
                handler(this, e);
            }
        }

        public event EventHandler<EventArgs> SelectedDestinationBranchChanged;
        private void OnSelectedDestinationBranchChanged(object sender, EventArgs e)
        {
            var handler = SelectedDestinationBranchChanged;
            if (handler != null)
            {
                handler(this, e);
            }
        }
    }
}
