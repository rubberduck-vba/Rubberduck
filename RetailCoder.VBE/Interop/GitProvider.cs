using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;
using Rubberduck.SourceControl;
using Microsoft.Vbe.Interop;
using System.Runtime.InteropServices;
using System.ComponentModel;

namespace Rubberduck.Interop
{
    [ComVisible(true)]
    [Guid("0C22A01D-3255-4BB6-8D67-DCC40A548A32")]
    [ProgId("Rubberduck.GitProvider")]
    [ClassInterface(ClassInterfaceType.None)]
    [System.ComponentModel.Description("VBA Editor integrated access to Git.")]
    class GitProvider : SourceControl.GitProvider, ISourceControlProvider
    {
        public GitProvider(VBProject project) 
            : base(project){}

        public GitProvider(VBProject project, IRepository repository)
            : base(project, repository){}

        public GitProvider(VBProject project, IRepository repository, string userName, string passWord)
            : base(project, repository, userName, passWord){}

        private IEnumerable branches;
        public new IEnumerable Branches
        {
            get { return new Branches(base.Branches); }
        }

        public new IFileStatusEntries Status()
        {
            return new FileStatusEntries(base.Status());
        }
    }
}
