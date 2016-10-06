using Rubberduck.SourceControl;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.SourceControl
{
    public interface ISourceControlProviderFactory
    {
        ISourceControlProvider CreateProvider(IVBProject project);
        ISourceControlProvider CreateProvider(IVBProject project, IRepository repository);
        ISourceControlProvider CreateProvider(IVBProject isAny, IRepository repository, SecureCredentials secureCredentials);
    }

    public class SourceControlProviderFactory : ISourceControlProviderFactory
    {
        public ISourceControlProvider CreateProvider(IVBProject project)
        {
            return new GitProvider(project);
        }

        public ISourceControlProvider CreateProvider(IVBProject project, IRepository repository)
        {
            return new GitProvider(project, repository);
        }

        public ISourceControlProvider CreateProvider(IVBProject project, IRepository repository, SecureCredentials creds)
        {
            return new GitProvider(project, repository, creds);
        }
    }
}
