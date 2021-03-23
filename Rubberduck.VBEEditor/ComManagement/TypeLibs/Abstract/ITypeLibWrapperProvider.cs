using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract
{
    public interface ITypeLibWrapperProvider : ITypeLibWrapperProviderLite
    {
        ITypeLibWrapper TypeLibWrapperFromProject(string projectId);
    }

    public interface ITypeLibWrapperProviderLite
    {
        ITypeLibWrapper TypeLibWrapperFromProject(IVBProject project);
    }
}
