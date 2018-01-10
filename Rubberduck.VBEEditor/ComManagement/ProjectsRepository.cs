using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
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

        private readonly ReaderWriterLockSlim _refreshProtectionLock = new ReaderWriterLockSlim(LockRecursionPolicy.NoRecursion);

        public ProjectsRepository(IVBE vbe)
        {
            _projectsCollection = vbe.VBProjects;
            LoadCollections();
        }

        private void LoadCollections()
        {
            LoadProjects();
            LoadComponentsCollections();
            LoadComponents();
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

        private void LoadComponents()
        {
            foreach (var components in _componentsCollections.Values)
            {
                LoadComponents(components);
            }
        }

        private void LoadComponents(IVBComponents componentsCollection)
        {
            foreach (var component in componentsCollection)
            {
                _components.Add(new QualifiedModuleName(component), component);
            }
        }

        public void Refresh()
        {
            throw new NotImplementedException();

            if (_disposed)
            {
                return;
            }

            ExecuteWithinWriteLock(() => RefreshCollections());
        }

        private void ExecuteWithinWriteLock(Action action)
        {
            var writeLockTaken = false;
            try
            {
                _refreshProtectionLock.EnterWriteLock();
                writeLockTaken = true;
                action.Invoke();
            }
            finally
            {
                if (writeLockTaken)
                {
                    _refreshProtectionLock.ExitWriteLock();
                }
            }
        }

        private void RefreshCollections()
        {
            //We save a copy of the collections and only refresh after the collections have been loaded again
            //to avoid disconnecting any RCWs from the underlying COM object for objects that still exist.
            var projects = ClearComWrapperDictionary(_projects);
            var componentCollections = ClearComWrapperDictionary(_componentsCollections);
            var components = ClearComWrapperDictionary(_components);

            LoadCollections();

            DisposeWrapperEnumerable(projects);
            DisposeWrapperEnumerable(componentCollections);
            DisposeWrapperEnumerable(components);
        }

        private IEnumerable<ISafeComWrapper> ClearComWrapperDictionary<TKey, TWrapper>(IDictionary<TKey, TWrapper> dictionary)
            where TWrapper : ISafeComWrapper
        {
            var copy = dictionary.Values.ToList() as IEnumerable<ISafeComWrapper>;
            dictionary.Clear();
            return copy;
        }

        private void DisposeWrapperEnumerable<TWrapper>(IEnumerable<TWrapper> wrappers) where TWrapper : ISafeComWrapper
        {
            foreach (var wrapper in wrappers)
            {
                wrapper.Dispose();
            }
        }

        private void RefreshCollections(string projectId)
        {
            IVBProject project;
            if (!_projects.TryGetValue(projectId, out project))
            {
                return;
            }

            var componentsCollection = _componentsCollections[projectId];
            var components = _components.Where(kvp => kvp.Key.ProjectId.Equals(projectId)).ToList();

            foreach (var qmn in components.Select(kvp => kvp.Key))
            {
                _components.Remove(qmn);
            }

            _componentsCollections[projectId] = project.VBComponents;
            LoadComponents(_componentsCollections[projectId]);

            componentsCollection.Dispose();
            DisposeWrapperEnumerable(components.Select(kvp => kvp.Value));
        }

        public void Refresh(string projectId)
        {
            throw new NotImplementedException();

            if (_disposed)
            {
                return;
            }

            ExecuteWithinWriteLock(() => RefreshCollections(projectId));
        }

        public IVBProjects ProjectsCollection()
        {
            throw new NotImplementedException();

            return _projectsCollection;
        }

        public IEnumerable<(string ProjectId, IVBProject Project)> Projects()
        {
            throw new NotImplementedException();

            return EvaluateWithinReadLock(() => _projects.Select(kvp => (kvp.Key, kvp.Value)).ToList());
        }

        private T EvaluateWithinReadLock<T>(Func<T> function)
        {
            var readLockTaken = false;
            try
            {
                _refreshProtectionLock.EnterReadLock();
                readLockTaken = true;
                return function.Invoke();
            }
            finally
            {
                if (readLockTaken)
                {
                    _refreshProtectionLock.ExitReadLock();
                }
            }
        }

        public IVBProject Project(string projectId)
        {
            throw new NotImplementedException();

            return EvaluateWithinReadLock(() => _projects.TryGetValue(projectId, out var project) ? project : null);
        }

        public IVBComponents ComponentsCollection(string projectId)
        {
            throw new NotImplementedException();

            return EvaluateWithinReadLock(() => _componentsCollections.TryGetValue(projectId, out var componenstCollection) ? componenstCollection : null);
        }

        public IEnumerable<(QualifiedModuleName QualifiedModuleName, IVBComponent Component)> Components()
        {
            throw new NotImplementedException();

            return EvaluateWithinReadLock(() => _components.Select(kvp => (kvp.Key, kvp.Value)).ToList());
        }

        public IEnumerable<(QualifiedModuleName QualifiedModuleName, IVBComponent Component)> Components(string projectId)
        {
            throw new NotImplementedException();

            return EvaluateWithinReadLock(() => _components.Where(kvp => kvp.Key.ProjectId.Equals(projectId))
                                                            .Select(kvp => (kvp.Key, kvp.Value))
                                                            .ToList());
        }

        public IVBComponent Component(QualifiedModuleName qualifiedModuleName)
        {
            return EvaluateWithinReadLock(() => _components.TryGetValue(qualifiedModuleName, out var component) ? component : null);
        }

        private bool _disposed;
        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }
            _disposed = true;

            ExecuteWithinWriteLock(() => ClearCollections());

            _refreshProtectionLock.Dispose();
        }

        private void ClearCollections()
        {
            var projects = ClearComWrapperDictionary(_projects);
            var componentCollections = ClearComWrapperDictionary(_componentsCollections);
            var components = ClearComWrapperDictionary(_components);

            DisposeWrapperEnumerable(projects);
            DisposeWrapperEnumerable(componentCollections);
            DisposeWrapperEnumerable(components);
        }
    }
}
