using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract
{
    public interface ITypeLibWrapperProvider
    {
        ITypeLibWrapper TypeLibWrapperFromProject(string projectId);
        ITypeLibWrapper TypeLibWrapperFromProject(IVBProject project);
    }
}
