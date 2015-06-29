using Microsoft.Vbe.Interop;
using Rubberduck.SourceControl;

namespace Rubberduck.UI.SourceControl
{
    public interface ISourceControlProviderFactory
    {
        ISourceControlProvider CreateProvider(VBProject project);
        ISourceControlProvider CreateProvider(VBProject project, IRepository repository);
        ISourceControlProvider CreateProvider(VBProject isAny, IRepository repository, SecureCredentials secureCredentials);
    }

    public class SourceControlProviderFactory : ISourceControlProviderFactory
    {
        public ISourceControlProvider CreateProvider(VBProject project)
        {
            return new GitProvider(project);
        }

        public ISourceControlProvider CreateProvider(VBProject project, IRepository repository)
        {
            return new GitProvider(project, repository);
        }

        public ISourceControlProvider CreateProvider(VBProject project, IRepository repository, SecureCredentials creds)
        {
            return new GitProvider(project, repository, creds);
        }
    }
}
