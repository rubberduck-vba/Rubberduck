namespace Rubberduck.VBEditor.ComManagement
{
    public interface IProjectsRepository : IProjectsProvider
    {
        void Refresh();
        void Refresh(string projectId);
        void RemoveComponent(QualifiedModuleName qualifiedModuleName);
    }
}
