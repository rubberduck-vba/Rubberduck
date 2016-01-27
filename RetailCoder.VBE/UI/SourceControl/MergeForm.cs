using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Windows.Forms;

namespace Rubberduck.UI.SourceControl
{
    [ExcludeFromCodeCoverage]
    public partial class MergeForm : Form, IMergeView
    {
        public MergeForm()
        {
            InitializeComponent();

            Text = RubberduckUI.SourceControl_MergeFormCaption;
            TitleLabel.Text = RubberduckUI.SourceControl_MergeFormTitle;
            InstructionsLabel.Text = RubberduckUI.SourceControl_MergeFormInstructions;
            SourceLabel.Text = RubberduckUI.SourceControl_SourceLabel;
            DestinationLabel.Text = RubberduckUI.SourceControl_DestinationLabel;
            OkButton.Text = RubberduckUI.OK;
            OkButton.Click += OnConfirm;
            CancelButton.Text = RubberduckUI.CancelButtonText;
            CancelButton.Click += OnCancel;
        }

        public bool OkButtonEnabled
        {
            get { return OkButton.Enabled; }
            set { OkButton.Enabled = value; }
        }

        private BindingList<string> _source;
        public IList<string> SourceSelectorData
        {
            get { return _source; }
            set
            {
                _source = new BindingList<string>(value);
                SourceSelector.DataSource = _source;
            }
        }

        private BindingList<string> _destination;
        public IList<string> DestinationSelectorData
        {
            get { return _destination; }
            set
            {
                _destination = new BindingList<string>(value);
                DestinationSelector.DataSource = _destination;
            }
        }

        public string SelectedSourceBranch
        {
            get { return SourceSelector.SelectedItem.ToString(); }
            set { SourceSelector.SelectedItem = value; }
        }

        public string SelectedDestinationBranch
        {
            get { return DestinationSelector.SelectedItem.ToString(); }
            set { DestinationSelector.SelectedItem = value; }
        }

        public string StatusText
        {
            get { return StatusTextBox.Text; }
            set { StatusTextBox.Text = value; }
        }

        public bool StatusTextVisible
        {
            get { return StatusTextBox.Visible; }
            set { StatusTextBox.Visible = value; }
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
