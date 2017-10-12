using Rubberduck.SourceControl;
using Rubberduck.VBEditor.SafeComWrappers.VB.Abstract;

namespace Rubberduck.UI.SourceControl
{
    public interface ISourceControlProviderFactory
    {
        ISourceControlProvider CreateProvider(IVBProject project);
        ISourceControlProvider CreateProvider(IVBProject project, IRepository repository);
        ISourceControlProvider CreateProvider(IVBProject project, IRepository repository, SecureCredentials secureCredentials);
    }
}
