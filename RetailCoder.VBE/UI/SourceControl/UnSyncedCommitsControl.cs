using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Windows.Forms;
using Rubberduck.SourceControl;

namespace Rubberduck.UI.SourceControl
{
    [ExcludeFromCodeCoverage]
    public partial class UnsyncedCommitsControl : UserControl, IUnsyncedCommitsView
    {
        public UnsyncedCommitsControl()
        {
            InitializeComponent();
            SetText();
        }

        public string CurrentBranch 
        { 
            get { return this.UnsyncedCommitsBranchNameLabel.Text; }
            set { this.UnsyncedCommitsBranchNameLabel.Text = value; }
        }

        private BindingList<ICommit> _incomingCommits;
        public IList<ICommit> IncomingCommits
        {
            get { return _incomingCommits; }
            set
            {
               _incomingCommits = new BindingList<ICommit>(value);
                this.IncomingCommitsGrid.DataSource = _incomingCommits;
            }
        }

        private BindingList<ICommit> _outgoingCommits;
        public IList<ICommit> OutgoingCommits
        {
            get { return _outgoingCommits; }
            set
            {
                _outgoingCommits = new BindingList<ICommit>(value);
                this.OutgoingCommitsGrid.DataSource = _outgoingCommits;
            }
        }

        public event EventHandler<EventArgs> Fetch;
        private void FetchIncomingCommitsButton_Click(object sender, EventArgs e)
        {
            var handler = Fetch;
            if (handler != null)
            {
                handler(this, e);
            }
        }

        public event EventHandler<EventArgs> Pull;
        private void PullButton_Click(object sender, EventArgs e)
        {
            var handler = Pull;
            if (handler != null)
            {
                handler(this, e);
            }
        }

        public event EventHandler<EventArgs> Push;
        private void PushButton_Click(object sender, EventArgs e)
        {
            var handler = Push;
            if (handler != null)
            {
                handler(this, e);
            }
        }

        public event EventHandler<EventArgs> Sync;
        private void SyncButton_Click(object sender, EventArgs e)
        {
            var handler = Sync;
            if (handler != null)
            {
                handler(this, e);
            }
        }

        private void SetText()
        {
            CurrentBranchLabel.Text = RubberduckUI.SourceControl_CurrentBranchLabel;
            FetchIncomingCommitsButton.Text = RubberduckUI.SourceControl_FetchCommitsLabel;
            PullButton.Text = RubberduckUI.SourceControl_PullCommitsLabel;
            PushButton.Text = RubberduckUI.SourceControl_PushCommitsLabel;
            SyncButton.Text = RubberduckUI.SourceControl_SyncCommitsLabel;

            IncomingCommitsBox.Text = RubberduckUI.SourceControl_IncomingCommits;
            OutgoingCommitsBox.Text = RubberduckUI.SourceControl_OutgoingCommits;
        }
    }
}
