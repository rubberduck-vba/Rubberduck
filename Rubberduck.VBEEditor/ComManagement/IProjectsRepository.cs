namespace Rubberduck.VBEditor.ComManagement
{
    public interface IProjectsRepository : IProjectsProvider
    {
        void Refresh(string projectId = null);
    }
}
