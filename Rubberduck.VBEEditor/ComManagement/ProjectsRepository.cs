using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.VBEditor.ComManagement
{
    public class ProjectsRepository : IProjectsRepository, IDisposable
    {
        private readonly IVBProjects _projectsCollection;
        private readonly IDictionary<string, IVBProject> _projects = new Dictionary<string, IVBProject>();
        private readonly IDictionary<string, IVBComponents> _componentsCollections = new Dictionary<string, IVBComponents>();
        private readonly IDictionary<QualifiedModuleName, IVBComponent> _components = new Dictionary<QualifiedModuleName, IVBComponent>();

        public ProjectsRepository(IVBE vbe)
        {
            _projectsCollection = vbe.VBProjects;
            LoadCollections();
        }

        private void LoadCollections()
        {
            LoadProjects();
            LoadComponentsCollections();
        }

        private void LoadProjects()
        {
            foreach (var project in _projectsCollection)
            {
                _projects.Add(project.ProjectId, project);
            }
        }

        private void LoadComponentsCollections()
        {
            foreach (var (projectId, project) in _projects)
            {
                _componentsCollections.Add(projectId, project.VBComponents);
            }
        }


        public IVBProjects ProjectsCollection()
        {
            throw new NotImplementedException();
        }

        public IEnumerable<(string ProjectId, IVBProject Project)> Projects()
        {
            throw new NotImplementedException();
        }

        public IVBProject Project(string projectId)
        {
            throw new NotImplementedException();
        }

        public IVBComponents ComponentsCollection(string projectId)
        {
            throw new NotImplementedException();
        }

        public IEnumerable<(QualifiedModuleName QualifiedModuleName, IVBComponent Component)> Components(string projectId = null)
        {
            throw new NotImplementedException();
        }

        public IVBComponent Component(QualifiedModuleName qualifiedModuleName)
        {
            throw new NotImplementedException();
        }

        public void Refresh(string projectId = null)
        {
            throw new NotImplementedException();
        }

        private bool _disposed;
        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }
            _disposed = true;

            
        }
    }
}
