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
            Repository repo = new Repository("SourceControlTest", @"C:\Users\Christopher\Documents\SourceControlTest", @"https://github.com/ckuhn203/SourceControlTest.git");
            this.git = new GitProvider(project, repo, "ckuhn203", "Macc2232");
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
    }
}
