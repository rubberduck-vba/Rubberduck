using System.Collections.Generic;
using System.Linq;
using NLog;
using Rubberduck.InternalApi.Extensions;
using Rubberduck.Parsing.Common;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA.Extensions;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.ComReflection
{
    public class UserProjectRepository : IUserComProjectRepository
    {

        private readonly IUiDispatcher _uiDispatcher;
        private readonly ITypeLibWrapperProvider _typeLibWrapperProvider;
        private readonly IProjectsProvider _projectsProvider;

        private readonly IDictionary<string, ComProject> _userComProjects = new Dictionary<string, ComProject>();

        private readonly Logger _logger = LogManager.GetCurrentClassLogger();


        public UserProjectRepository(ITypeLibWrapperProvider typeLibWrapperProvider, IUiDispatcher uiDispatcher,
            IProjectsProvider projectsProvider)
        {
            _typeLibWrapperProvider = typeLibWrapperProvider;
            _uiDispatcher = uiDispatcher;
            _projectsProvider = projectsProvider;
        }

        //NOTE: Before this class can be used outside the parsing process thread-safety measures will have to be added.
        public ComProject UserProject(string projectId)
        {
            return _userComProjects.TryGetValue(projectId, out var result) ? result : null;
        }

        public IReadOnlyDictionary<string, ComProject> UserProjects()
        {
            //We return a copy to avoid issues updating the collections contents. 
            return _userComProjects.ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
        }

        public void RefreshUserComProjects(IReadOnlyCollection<string> projectIdsToReload)
        {
            var parsingStageTimer = ParsingStageTimer.StartNew();

            RemoveNoLongerExistingProjects();
            RemoveProjects(projectIdsToReload);
            var loadTask = _uiDispatcher.StartTask(() =>
            {
                AddUnprotectedUserProjects(projectIdsToReload);
                AddLockedProjects();
            });
            loadTask.Wait();

            parsingStageTimer.Stop();
            parsingStageTimer.Log("Loaded ComProjects for user projects in {0}ms.");
        }

        private void RemoveNoLongerExistingProjects()
        {
            var existingProjectIds = _projectsProvider.Projects().Select(tpl => tpl.ProjectId)
                .Concat(_projectsProvider.LockedProjects().Select(tpl => tpl.ProjectId)).ToHashSet();
            var noLongerExistingProjectIds = _userComProjects.Keys
                .Where(projectId => !existingProjectIds.Contains(projectId))
                .ToList();
            RemoveProjects(noLongerExistingProjectIds);
        }

        /// <summary>
        /// Use only on the UI thread!
        /// </summary>
        /// <remarks>
        ///This method uses TryLoadProject, which is only safe to use on the UI thread.
        /// </remarks>
        private void AddLockedProjects()
        {
            var lockedProjects = _projectsProvider.LockedProjects();
            foreach (var (projectId, project) in lockedProjects)
            {
                if (_userComProjects.ContainsKey(projectId))
                {
                    continue;
                }

                if (TryLoadProject(projectId, project, out var comProject))
                {
                    _userComProjects.Add(projectId, comProject);
                } 
            }
        }

        /// <summary>
        /// Use only on the UI thread!
        /// </summary>
        /// <remarks>
        ///This method uses the typeLib API, which is only safe to use on the UI thread.
        /// </remarks>
        private bool TryLoadProject(string projectId, IVBProject project, out ComProject comProject)
        {
            if (!project.TryGetFullPath(out var path))
            {
                //Only debug because this will always happen for unsaved projects.
                _logger.Debug($"Unable to get project path for project with projectId {projectId} when loading the COM project.");
                path = string.Empty;
            }

            using (var typeLib = _typeLibWrapperProvider.TypeLibWrapperFromProject(project))
            {
                comProject = typeLib != null 
                    ? new ComProject(typeLib, path) 
                    : null;
            }

            return comProject != null;
        } 

        private void RemoveProjects(IEnumerable<string> projectIdsToRemove)
        {
            foreach (var projectId in projectIdsToRemove)
            {
                _userComProjects.Remove(projectId);
            }
        }

        /// <summary>
        /// Use only on the UI thread!
        /// </summary>
        /// <remarks>
        ///This method uses TryLoadProject, which is only safe to use on the UI thread.
        /// </remarks>
        private void AddUnprotectedUserProjects(IReadOnlyCollection<string> projectIdsToLoad)
        {
            var projectsToLoad = _projectsProvider.Projects().Where(tpl => projectIdsToLoad.Contains(tpl.ProjectId));
            foreach (var (projectId, project) in projectsToLoad)
            {
                if (TryLoadProject(projectId, project, out var comProject))
                {
                    _userComProjects.Add(projectId, comProject);
                }
            }
        }
    }
}
