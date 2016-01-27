using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Windows.Forms;

namespace Rubberduck.UI.SourceControl
{
    [ExcludeFromCodeCoverage]
    public partial class DeleteBranchForm : Form, IDeleteBranchView
    {
        public DeleteBranchForm()
        {
            InitializeComponent();

            Branches = new List<string>();

            Text = RubberduckUI.SourceControl_DeleteBranchCaption;
            TitleLabel.Text = RubberduckUI.SourceControl_DeleteBranchTitleLable;
            InstructionsLabel.Text = RubberduckUI.SourceControl_DeleteBranchInstructionsLabel;
            OkButton.Text = RubberduckUI.OK_AllCaps;

            OkButton.Click += OkButton_Click;
            CancelButton.Text = RubberduckUI.CancelButtonText;
            CancelButton.Click += CancelButton_Click;
            BranchList.SelectedValueChanged += BranchList_SelectedValueChanged;
        }

        public event EventHandler<BranchDeleteArgs> SelectionChanged;
        private void BranchList_SelectedValueChanged(object sender, EventArgs e)
        {
            var handler = SelectionChanged;
            if (handler != null)
            {
                handler(this, new BranchDeleteArgs(BranchList.SelectedItem.ToString()));
            }
        }

        public bool OkButtonEnabled
        {
            get { return OkButton.Enabled; }
            set { OkButton.Enabled = value; }
        }

        private IList<string> _branches;
        public IList<string> Branches
        {
            get { return _branches; }
            set
            {
                _branches = value;
                BranchList.DataSource = Branches;
                BranchList.Refresh();
           }
        }

        public event EventHandler<BranchDeleteArgs> Confirm;
        private void OkButton_Click(object sender, EventArgs e)
        {
            var handler = Confirm;
            if (handler != null)
            {
                handler(this, new BranchDeleteArgs(BranchList.SelectedItem.ToString()));
            }
        }

        public event EventHandler<EventArgs> Cancel;
        private void CancelButton_Click(object sender, EventArgs e)
        {
            var handler = Cancel;
            if (handler != null)
            {
                handler(this, e);
            }
        }
    }
}
