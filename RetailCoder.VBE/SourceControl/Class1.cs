using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LibGit2Sharp;
using LibGit2Sharp.Core;

namespace Rubberduck.SourceControl
{
    class SourceControlSandbox
    {

        public void sandbox()
        {
            Repository repo;
            LibGit2Sharp.MergeResult result;
            
            //repo.CreateBranch()
            //repo.Branches.Where(x => x.IsRemote == false);
            ISourceControlProvider git = new GitProvider(new Microsoft.Vbe.Interop.VBProject(),
                                        new Repository("SourceControlTest", @"C:\Users\Christopher\Documents\SourceControlTest", @"https://github.com/ckuhn203/SourceControlTest.git")
                                        );
            
        }
    }
}
