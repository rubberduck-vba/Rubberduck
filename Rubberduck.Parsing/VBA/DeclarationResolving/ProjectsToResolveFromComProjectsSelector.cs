using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBEditor.ComManagement;

namespace Rubberduck.Parsing.VBA.DeclarationResolving
{
    public class ProjectsToResolveFromComProjectsSelector : IProjectsToResolveFromComProjectSelector
    {
        private readonly IProjectsProvider _projectsProvider;
        private readonly HashSet<string> _projectsToResolveFromComProjects = new HashSet<string>();

        private readonly object _collectionLockObject = new object();

        public ProjectsToResolveFromComProjectsSelector(IProjectsProvider projectsProvider)
        {
            _projectsProvider = projectsProvider;
        }


        public IReadOnlyCollection<string> ProjectsToResolveFromComProjects {
            get
            {
                lock (_collectionLockObject)
                {
                    return _projectsToResolveFromComProjects;
                }
            }
        }

        public void RefreshProjectsToResolveFromComProjectSelector()
        {
            lock (_collectionLockObject)
            {
                _projectsToResolveFromComProjects.Clear();
                _projectsToResolveFromComProjects.UnionWith(_projectsProvider.LockedProjects().Select(tpl => tpl.ProjectId));
            }
        }

        public bool ToBeResolvedFromComProject(string projectId)
        {
            lock (_collectionLockObject)
            {
                return _projectsToResolveFromComProjects.Contains(projectId);
            }
        }
    }
}