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
    public partial class GitView : Form
    {
        private ISourceControlProvider git;

        public GitView()
        {
            InitializeComponent();
        }

        public GitView(Microsoft.Vbe.Interop.VBProject project):this()
        {
            Repository repo = new Repository("SourceControlTest", @"path", @"url");
            this.git = new GitProvider(project, repo, "username", "password");
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
            git.Checkout("master");
        }
    }
}
