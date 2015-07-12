using Microsoft.Vbe.Interop;
using Rubberduck.SourceControl;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.UI.SourceControl
{
    public interface ISourceControlProviderFactory
    {
        ISourceControlProvider CreateProvider(VBProject project);
        ISourceControlProvider CreateProvider(VBProject project, IRepository repository, IRubberduckCodePaneFactory factory);
        ISourceControlProvider CreateProvider(VBProject isAny, IRepository repository, SecureCredentials secureCredentials, IRubberduckCodePaneFactory factory);
    }

    public class SourceControlProviderFactory : ISourceControlProviderFactory
    {
        public ISourceControlProvider CreateProvider(VBProject project)
        {
            return new GitProvider(project);
        }

        public ISourceControlProvider CreateProvider(VBProject project, IRepository repository, IRubberduckCodePaneFactory factory)
        {
            return new GitProvider(project, repository, factory);
        }

        public ISourceControlProvider CreateProvider(VBProject project, IRepository repository, SecureCredentials creds, IRubberduckCodePaneFactory factory)
        {
            return new GitProvider(project, repository, creds, factory);
        }
    }
}
