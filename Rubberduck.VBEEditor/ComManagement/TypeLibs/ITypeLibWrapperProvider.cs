namespace Rubberduck.VBEditor.ComManagement.TypeLibs
{
    public interface ITypeLibWrapperProvider
    {
        ITypeLibWrapper TypeLibWrapperFromProject(string projectId);
    }
}
