using Microsoft.Vbe.Interop;
using Rubberduck.SourceControl;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.UI.SourceControl
{
    public interface ISourceControlProviderFactory
    {
        ISourceControlProvider CreateProvider(VBProject project);
        ISourceControlProvider CreateProvider(VBProject project, IRepository repository, ICodePaneWrapperFactory wrapperFactory);
        ISourceControlProvider CreateProvider(VBProject isAny, IRepository repository, SecureCredentials secureCredentials, ICodePaneWrapperFactory wrapperFactory);
    }

    public class SourceControlProviderFactory : ISourceControlProviderFactory
    {
        public ISourceControlProvider CreateProvider(VBProject project)
        {
            return new GitProvider(project);
        }

        public ISourceControlProvider CreateProvider(VBProject project, IRepository repository, ICodePaneWrapperFactory wrapperFactory)
        {
            return new GitProvider(project, repository, wrapperFactory);
        }

        public ISourceControlProvider CreateProvider(VBProject project, IRepository repository, SecureCredentials creds, ICodePaneWrapperFactory wrapperFactory)
        {
            return new GitProvider(project, repository, creds, wrapperFactory);
        }
    }
}
