using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
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

        public DummyGitView(Microsoft.Vbe.Interop.VBProject project):this()
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
    }
}
