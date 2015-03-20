using System;
using System.ComponentModel;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.SourceControl;

namespace Rubberduck.UI.SourceControl
{
    public partial class DummyGitView : Form
    {
        private ISourceControlProvider git;

        public DummyGitView()
        {
            InitializeComponent();
        }

        public DummyGitView(VBProject project):this()
        {
            Repository repo = new Repository("SourceControlTest", @"C:\Users\Christopher\Documents\SourceControlTest", @"https://github.com/ckuhn203/SourceControlTest.git");
            this.git = new GitProvider(project, repo, "UserName", "Password");
        }

        private void Commit_Click(object sender, EventArgs e)
        {
            git.Commit("TestCommit");
        }

        private void Push_Click(object sender, EventArgs e)
        {
            git.Push();
        }

        private void Pull_Click(object sender, EventArgs e)
        {
            git.Pull();
        }

        private void Fetch_Click(object sender, EventArgs e)
        {
            git.Fetch();
        }

        private void NewBranch_Click(object sender, EventArgs e)
        {
            git.CreateBranch("testbranch");
        }

        private void Checkout_Click(object sender, EventArgs e)
        {
            git.Checkout(this.SourceBranch.SelectedItem.ToString());
        }

        private void Merge_Click(object sender, EventArgs e)
        {
            git.Merge(this.SourceBranch.SelectedItem.ToString(), this.DestinationBranch.SelectedItem.ToString());
        }

        private void Undo_Click(object sender, EventArgs e)
        {
            git.Undo(@"C:\Users\Christopher\Documents\SourceControlTest\Module1.bas");
        }

        private void Revert_Click(object sender, EventArgs e)
        {
            git.Revert();
        }

        private void Status_Click(object sender, EventArgs e)
        {
            var status = git.Status();
            this.StatusResults.DataSource = new BindingList<IFileStatusEntry>(status.ToList());
        }
    }
}
