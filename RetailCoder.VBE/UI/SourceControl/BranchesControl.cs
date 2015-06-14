using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Windows.Forms;

namespace Rubberduck.UI.SourceControl
{
    [ExcludeFromCodeCoverage]
    public partial class BranchesControl : UserControl, IBranchesView
    {
        public BranchesControl()
        {
            InitializeComponent();

            SetText();
        }

        private void SetText()
        {
            CurrentBranch.Text = RubberduckUI.SourceControl_CurrentBranchLabel;
            NewBranchButton.Text = RubberduckUI.SourceControl_NewBranch;
            MergeBranchButton.Text = RubberduckUI.SourceControl_MergeBranch;
            PublishedBranchesBox.Text = RubberduckUI.SourceControl_PublishedBranchesLabel;
            UnpublishedBranchesBox.Text = RubberduckUI.SourceControl_UnpublishedBranchesLabel;
        }

        private BindingList<string> _branches;
        public IList<string> Local
        {
            get { return _branches; }
            set
            {
                _branches = new BindingList<string>(value);
                this.CurrentBranchSelector.DataSource = _branches;
            }
        }

        public string Current
        {
            get { return this.CurrentBranchSelector.SelectedItem.ToString(); }
            set { this.CurrentBranchSelector.SelectedItem = value; }
        }

        private BindingList<string> _publishedBranches;
        public IList<string> Published
        {
            get { return _publishedBranches; }
            set
            {
                _publishedBranches = new BindingList<string>(value);
                this.PublishedBranchesList.DataSource = _publishedBranches;
            }
        }

        private BindingList<string> _unpublishedBranches;
        public IList<string> Unpublished
        {
            get { return _unpublishedBranches; }
            set
            {
                _unpublishedBranches = new BindingList<string>(value);
                this.UnpublishedBranchesList.DataSource = _unpublishedBranches;
            }
        }

        public event EventHandler<EventArgs> SelectedBranchChanged;
        public void OnSelectedBranchChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(this.Current))
            {
                RaiseGenericEvent(SelectedBranchChanged, e);
            }
        }

        public event EventHandler<EventArgs> Merge;
        public void OnMerge(object sender, EventArgs e)
        {
            RaiseGenericEvent(Merge, e);
        }

        public event EventHandler<EventArgs> CreateBranch;
        public void OnCreateBranch(object sender, EventArgs e)
        {
            RaiseGenericEvent(CreateBranch, e);
        }

        private void RaiseGenericEvent(EventHandler<EventArgs> handler, EventArgs e)
        {
            if (handler != null)
            {
                handler(this, e);
            }
        }
    }
}
