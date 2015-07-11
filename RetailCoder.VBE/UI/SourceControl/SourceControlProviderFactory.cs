using Microsoft.Vbe.Interop;
using Rubberduck.SourceControl;
using Rubberduck.VBEditor.VBEInterfaces;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.UI.SourceControl
{
    public interface ISourceControlProviderFactory
    {
        ISourceControlProvider CreateProvider(VBProject project);
        ISourceControlProvider CreateProvider(VBProject project, IRepository repository, IRubberduckFactory<IRubberduckCodePane> factory);
        ISourceControlProvider CreateProvider(VBProject isAny, IRepository repository, SecureCredentials secureCredentials, IRubberduckFactory<IRubberduckCodePane> factory);
    }

    public class SourceControlProviderFactory : ISourceControlProviderFactory
    {
        public ISourceControlProvider CreateProvider(VBProject project)
        {
            return new GitProvider(project);
        }

        public ISourceControlProvider CreateProvider(VBProject project, IRepository repository, IRubberduckFactory<IRubberduckCodePane> factory)
        {
            return new GitProvider(project, repository, factory);
        }

        public ISourceControlProvider CreateProvider(VBProject project, IRepository repository, SecureCredentials creds, IRubberduckFactory<IRubberduckCodePane> factory)
        {
            return new GitProvider(project, repository, creds, factory);
        }
    }
}
