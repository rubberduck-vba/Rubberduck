using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.VBA
{
    public class RepositoryProjectManager : IProjectManager
    {
        private readonly IProjectsRepository _projectsRepository;

        public RepositoryProjectManager(IProjectsRepository projectsRepository)
        {
            if (projectsRepository == null)
            {
                throw new ArgumentNullException(nameof(projectsRepository));
            }

            _projectsRepository = projectsRepository;
        }

        public IReadOnlyCollection<IVBProject> Projects => _projectsRepository.Projects().Select(tpl => tpl.Project).ToList().AsReadOnly();

        public void RefreshProjects()
        {
            _projectsRepository.Refresh();
        }

        public IReadOnlyCollection<QualifiedModuleName> AllModules()
        {
            return _projectsRepository.Components().Select(tpl => tpl.QualifiedModuleName).ToHashSet().AsReadOnly();
        }
    }
}
