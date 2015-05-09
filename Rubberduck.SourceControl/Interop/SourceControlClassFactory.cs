using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;

namespace Rubberduck.SourceControl.Interop
{
    [ComVisible(true)]
    [Guid("335DA0D8-625C-4CB9-90CD-C9A306B9B787")]
    public interface _ISourceControlClassFactory
    {
        [DispId(1)]
        ISourceControlProvider CreateGitProvider(VBProject project, [Optional] IRepository repository, [Optional] string userName, [Optional] string passWord);

        [DispId(2)]
        IRepository CreateRepository(string name, string localDirectory, [Optional] string remotePathOrUrl);
    }

    [ComVisible(true)]
    [Guid("29FB0A0E-F113-458F-823B-1CD1B60D2CA7")]
    [ProgId("Rubberduck.SourceControlClassFactory")]
    [ClassInterface(ClassInterfaceType.None)]
    public class SourceControlClassFactory : _ISourceControlClassFactory
    {
        [Description("Returns a new GitProvider. IRepository must be supplied if also passing user credentials.")]
        public ISourceControlProvider CreateGitProvider(VBProject project, [Optional] IRepository repository, [Optional] string userName, [Optional] string passWord)
        {
            if (passWord != null && userName != null)
            {
                if (repository == null)
                {
                    throw new ArgumentNullException("Must supply an IRepository if supplying credentials.");
                }
                return new GitProvider(project, repository, userName, passWord);
            }

            if (repository != null) 
            {
                return new GitProvider(project, repository);
            }

            return new GitProvider(project);
        }

        [Description("Returns new instance of repository struct.")]
        public IRepository CreateRepository(string name, string localDirectory, [Optional] string remotePathOrUrl)
        {
            if (remotePathOrUrl == null)
            {
                remotePathOrUrl = string.Empty;
            }

            return new Repository(name, localDirectory, remotePathOrUrl);
        }
    }
}
