using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.Parsing.VBA.DeclarationResolving
{
    public class ProjectsToResolveFromComProjectsSelector : IProjectsToResolveFromComProjectSelector
    {
        private readonly IProjectsProvider _projectsProvider;
        private readonly HashSet<string> _projectsToResolveFromComProjects = new HashSet<string>();
        private readonly IConfigurationService<IgnoredProjectsSettings> _ignoredProjectsSettingsProvider;

        private readonly object _collectionLockObject = new object();

        public ProjectsToResolveFromComProjectsSelector(IProjectsProvider projectsProvider, IConfigurationService<IgnoredProjectsSettings> ignoredProjectsSettingsProvider)
        {
            _projectsProvider = projectsProvider;
            _ignoredProjectsSettingsProvider = ignoredProjectsSettingsProvider;
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
                _projectsToResolveFromComProjects.UnionWith(LockedProjectIds());
                _projectsToResolveFromComProjects.UnionWith(IgnoredProjectIds());
            }
        }

        private IEnumerable<string> LockedProjectIds()
        {
            return _projectsProvider
                .LockedProjects()
                .Select(tpl => tpl.ProjectId);
        }

        private IEnumerable<string> IgnoredProjectIds()
        {
            var ignoredProjectFilenames = _ignoredProjectsSettingsProvider.Read().IgnoredProjectPaths;
            return _projectsProvider
                .Projects()
                .Where(tpl => tpl.Project.TryGetFullPath(out var filename) 
                                                && ignoredProjectFilenames.Contains(filename))
                .Select(tpl => tpl.ProjectId);
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