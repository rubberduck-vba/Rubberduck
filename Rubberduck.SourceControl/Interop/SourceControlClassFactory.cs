using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.SafeComWrappers.VBA;

namespace Rubberduck.SourceControl.Interop
{
    [ComVisible(true)]
    [Guid("335DA0D8-625C-4CB9-90CD-C9A306B9B787")]
    // ReSharper disable once InconsistentNaming; the underscore hides the interface from VBE's object browswer
    public interface _ISourceControlClassFactory
    {
        [DispId(1)]
        ISourceControlProvider CreateGitProvider(VBProject project, [Optional] IRepository repository, [Optional] ICredentials credentials);

        [DispId(2)]
        IRepository CreateRepository(string name, string localDirectory, [Optional] string remotePathOrUrl);

        [DispId(3)]
        ICredentials CreateCredentials(string username, string password);
    }

    [ComVisible(true)]
    [Guid("29FB0A0E-F113-458F-823B-1CD1B60D2CA7")]
    [ProgId("Rubberduck.SourceControlClassFactory")]
    [ClassInterface(ClassInterfaceType.None)]
    public class SourceControlClassFactory : _ISourceControlClassFactory
    {
        [Description("Returns a new GitProvider. IRepository must be supplied if also passing user credentials.")]
        public ISourceControlProvider CreateGitProvider(VBProject project, [Optional] IRepository repository, [Optional] ICredentials credentials)
        {
            if (credentials != null)
            {
                if (repository == null)
                {
                    throw new ArgumentNullException("Must supply an IRepository if supplying credentials.");
                }

                return new GitProvider(project, repository, credentials);
            }

            if (repository != null) 
            {
                return new GitProvider(project, repository);
            }

            return new GitProvider(project);
        }

        [Description("Returns new instance of type IRepository.")]
        public IRepository CreateRepository(string name, string localDirectory, [Optional] string remotePathOrUrl)
        {
            if (remotePathOrUrl == null)
            {
                remotePathOrUrl = string.Empty;
            }

            return new Repository(name, localDirectory, remotePathOrUrl);
        }

        [Description("Returns a new instance of type ICredentials.")]
        public ICredentials CreateCredentials(string username, string password)
        {
            return new Credentials(username, password);
        }
    }
}
