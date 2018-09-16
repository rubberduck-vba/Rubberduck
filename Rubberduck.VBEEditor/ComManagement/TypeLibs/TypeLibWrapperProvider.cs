using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs
{
    public class TypeLibWrapperProvider : ITypeLibWrapperProvider
    {
        private readonly IProjectsProvider _projectsProvider;

        public TypeLibWrapperProvider(IProjectsProvider projectsProvider)
        {
            _projectsProvider = projectsProvider;
        }

        public ITypeLibWrapper TypeLibWrapperFromProject(string projectId)
        {
            var project = _projectsProvider.Project(projectId);
            return TypeLibWrapperFromProject(project);
        }

        public ITypeLibWrapper TypeLibWrapperFromProject(IVBProject project)
        {
            return project != null ? TypeLibWrapper.FromVBProject(project) : null;
        }
    }
}