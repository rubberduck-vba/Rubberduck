namespace Rubberduck.VBEditor.ComManagement.TypeLibs
{
    public interface ITypeLibWrapperProvider
    {
        TypeLibWrapper TypeLibWrapperFromProject(string projectId);
    }
}
