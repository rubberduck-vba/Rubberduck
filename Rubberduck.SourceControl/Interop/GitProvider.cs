using System;
using System.Collections;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.SourceControl.Interop
{
    [ComVisible(true)]
    [Guid("0C22A01D-3255-4BB6-8D67-DCC40A548A32")]
    [ProgId("Rubberduck.GitProvider")]
    [ClassInterface(ClassInterfaceType.None)]
    [Description("VBA Editor integrated access to Git.")]
    class GitProvider : SourceControl.GitProvider, ISourceControlProvider
    {
        public GitProvider(VBProject project) 
            : base(project)
        { }

        public GitProvider(VBProject project, IRepository repository, IRubberduckCodePaneFactory factory)
            : base(project, repository, factory)
        { }

        [Obsolete]
        public GitProvider(VBProject project, IRepository repository, string userName, string passWord, IRubberduckCodePaneFactory factory)
            : base(project, repository, userName, passWord, factory)
        { }

        public GitProvider(VBProject project, IRepository repository, ICredentials credentials, IRubberduckCodePaneFactory factory)
            :base(project, repository, credentials.Username, credentials.Password, factory)
        { }

        public new string CurrentBranch
        {
            get { return base.CurrentBranch.Name; }
        }

        /// <summary>
        /// Returns only local branches to COM clients.
        /// </summary>
        public new IEnumerable Branches
        {
            get { return new Branches(base.Branches.Where(b => !b.IsRemote)); }
        }

        /// <summary>
        /// Returns Iterable Collection of FileStatusEntry objects.
        /// </summary>
        /// <returns></returns>
        public new IEnumerable Status()
        {
            return new FileStatusEntries(base.Status());
        }

        /// <summary>
        /// Stages and commits all modified files.
        /// </summary>
        /// <param name="message"></param>
        public override void Commit(string message)
        {
            var filePaths = base.Status()
                .Where(s => s.FileStatus.HasFlag(FileStatus.Modified))
                .Select(s => s.FilePath).ToList();

            Stage(filePaths);
            base.Commit(message);
        }
    }
}
