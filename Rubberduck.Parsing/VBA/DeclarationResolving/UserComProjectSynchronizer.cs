using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Common;
using Rubberduck.Parsing.ComReflection;

namespace Rubberduck.Parsing.VBA.DeclarationResolving
{
    public class UserComProjectSynchronizer : IUserComProjectSynchronizer
    {
        private readonly IDeclarationsFromComProjectLoader _declarationsFromComProjectLoader;
        private readonly IUserComProjectProvider _userComProjectProvider;
        private readonly IProjectsToResolveFromComProjectSelector _projectsToResolveFromComProjectSelector;

        private readonly RubberduckParserState _state;

        private readonly HashSet<string> _unloadedProjectIds = new HashSet<string>();
        private bool _lastSyncLoadedDeclaration; 

        private readonly HashSet<string> _currentlyLoadedProjectIds = new HashSet<string>();

        public UserComProjectSynchronizer(
            RubberduckParserState state,
            IDeclarationsFromComProjectLoader declarationsFromComProjectLoader,
            IUserComProjectProvider userComProjectProvider,
            IProjectsToResolveFromComProjectSelector projectsToResolveFromComProjectSelector)
        {
            _declarationsFromComProjectLoader = declarationsFromComProjectLoader;
            _userComProjectProvider = userComProjectProvider;
            _projectsToResolveFromComProjectSelector = projectsToResolveFromComProjectSelector;
            _state = state;
        }


        public bool LastSyncOfUserComProjectsLoadedDeclarations => _lastSyncLoadedDeclaration;
        public IReadOnlyCollection<string> UserProjectIdsUnloaded => _unloadedProjectIds;

        public void SyncUserComProjects()
        {
            var parsingStateTimer = ParsingStageTimer.StartNew();

            _lastSyncLoadedDeclaration = false;
            _unloadedProjectIds.Clear();

            var projectIdsToBeLoaded = _projectsToResolveFromComProjectSelector.ProjectsToResolveFromComProjects;
            var newProjectIdsToBeLoaded =
                projectIdsToBeLoaded.Where(projectId => !_currentlyLoadedProjectIds.Contains(projectId)).ToList();
            var projectsToBeUnloaded =
                _currentlyLoadedProjectIds.Where(projectId => !projectIdsToBeLoaded.Contains(projectId)).ToList();

            LoadProjects(newProjectIdsToBeLoaded);
            UnloadProjects(projectsToBeUnloaded);

            parsingStateTimer.Stop();
            parsingStateTimer.Log("Loaded declarations from ComProjects for user projects in {0}ms.");
        }

        private void LoadProjects(IEnumerable<string> projectIds)
        {
            foreach (var projectId in projectIds)
            {
                var comProject = _userComProjectProvider.UserProject(projectId);
                if (comProject == null)
                {
                    continue;
                }

                var declarations = _declarationsFromComProjectLoader.LoadDeclarations(comProject, projectId);
                foreach (var declaration in declarations)
                {
                    _state.AddDeclaration(declaration);
                }

                _currentlyLoadedProjectIds.Add(projectId);
                _lastSyncLoadedDeclaration = true;
            }
        }

        private void UnloadProjects(IEnumerable<string> projectIds)
        {
            foreach (var projectId in projectIds)
            {
                _state.RemoveBuiltInDeclarations(projectId);

                _unloadedProjectIds.Add(projectId);
                _currentlyLoadedProjectIds.Remove(projectId);
            }
        }
    }
}