using System.Windows.Forms;

namespace Rubberduck.UI.SourceControl
{
    public partial class UnSyncedCommitsControl : UserControl, IUnSyncedCommitsView
    {
        public UnSyncedCommitsControl()
        {
            InitializeComponent();

            SetText();
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
