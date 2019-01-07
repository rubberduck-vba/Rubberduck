using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.Events
{
    public sealed class VBEEvents : IVBEEvents
    {
        private static VBEEvents _instance;
        private static readonly object Lock = new object();
        private readonly IVBProjects _projects;
        private readonly Dictionary<string, IVBComponents> _components;
        private readonly Dictionary<string, IReferences> _references;

        public static VBEEvents Initialize(IVBE vbe)
        {
            lock (Lock)
            {
                if (_instance == null)
                {
                    _instance = new VBEEvents(vbe);
                }
            }

            return _instance;
        }

        public static void Terminate()
        {
            lock (Lock)
            {
                if (_instance == null)
                {
                    return;
                }

                _instance.Dispose();
                _instance = null;
            }
        }

        private VBEEvents(IVBE vbe)
        {
            _components = new Dictionary<string, IVBComponents>();
            _references = new Dictionary<string, IReferences>();

            if (_projects != null)
            {
                return;
            }
            
            _projects = vbe.VBProjects;

            _projects.AttachEvents();
            _projects.ProjectAdded += ProjectAddedHandler;
            _projects.ProjectRemoved += ProjectRemovedHandler;
            _projects.ProjectRenamed += ProjectRenamedHandler;
            _projects.ProjectActivated += ProjectActivatedHandler;
            foreach (var project in _projects)
            using (project)
            {
                {
                    RegisterProjectHandlers(project);
                }
            }
        }

        private void RegisterProjectHandlers(string projectId, string projectName)
        {
            IVBProject project = null;
            foreach (var item in _projects)
            {
                if (item.ProjectId == projectId && item.Name == projectName)
                {
                    project = item;
                    break;
                }

                item.Dispose();
            }

            if (project == null)
            {
                return;
            }

            RegisterProjectHandlers(project);
        }

        private void RegisterProjectHandlers(IVBProject project)
        {
            if (project.IsWrappingNullReference || project.Protection != ProjectProtection.Unprotected)
            {
                return;
            }

            project.AssignProjectId();

            var components = project.VBComponents;
            _components.Add(project.ProjectId, components);

            components.AttachEvents();
            components.ComponentAdded += ComponentAddedHandler;
            components.ComponentRemoved += ComponentRemovedHandler;
            components.ComponentRenamed += ComponentRenamedHandler;
            components.ComponentActivated += ComponentActivatedHandler;
            components.ComponentSelected += ComponentSelectedHandler;
            components.ComponentReloaded += ComponentReloadedHandler;

            var references = project.References;
            _references.Add(project.ProjectId, references);
            references.AttachEvents();
            references.ItemAdded += ProjectReferenceAddedHandler;
            references.ItemRemoved += ProjectReferenceRemovedHandler;
        }

        private void UnregisterProjectHandlers(string projectId)
        {
            if (!_components.ContainsKey(projectId))
            {
                return;
            }

            using (var components = _components[projectId])
            {
                components.ComponentAdded -= ComponentAddedHandler;
                components.ComponentRemoved -= ComponentRemovedHandler;
                components.ComponentRenamed -= ComponentRenamedHandler;
                components.ComponentActivated -= ComponentActivatedHandler;
                components.ComponentSelected -= ComponentSelectedHandler;
                components.ComponentReloaded -= ComponentReloadedHandler;
                components.DetachEvents();

                _components.Remove(projectId);
            }

            using (var references = _references[projectId])
            {
                references.ItemAdded -= ProjectReferenceAddedHandler;
                references.ItemRemoved -= ProjectReferenceRemovedHandler;
                references.DetachEvents();
                _references.Remove(projectId);
            }
        }

        public event EventHandler<ProjectEventArgs> ProjectAdded;
        private void ProjectAddedHandler(object sender, ProjectEventArgs e)
        {
            if (!_components.ContainsKey(e.ProjectId))
            {
                RegisterProjectHandlers(e.ProjectId, e.ProjectName);
            }
            ProjectAdded?.Invoke(sender, e);
        }

        public event EventHandler<ProjectEventArgs> ProjectRemoved;
        private void ProjectRemovedHandler(object sender, ProjectEventArgs e)
        {
            UnregisterProjectHandlers(e.ProjectId);
            ProjectRemoved?.Invoke(sender, e);
        }

        public event EventHandler<ProjectRenamedEventArgs> ProjectRenamed; 
        private void ProjectRenamedHandler(object sender, ProjectRenamedEventArgs e)
        {
            ProjectRenamed?.Invoke(sender, e);
        }

        public event EventHandler<ProjectEventArgs> ProjectActivated; 
        private void ProjectActivatedHandler(object sender, ProjectEventArgs e)
        {
            ProjectActivated?.Invoke(sender, e);
        }

        public event EventHandler<ComponentEventArgs> ComponentAdded;
        private void ComponentAddedHandler(object sender, ComponentEventArgs e)
        {
            ComponentAdded?.Invoke(sender, e);
        }

        public event EventHandler<ComponentEventArgs> ComponentRemoved; 
        private void ComponentRemovedHandler(object sender, ComponentEventArgs e)
        {
            ComponentRemoved?.Invoke(sender, e);
        }

        public event EventHandler<ComponentRenamedEventArgs> ComponentRenamed; 
        private void ComponentRenamedHandler(object sender, ComponentRenamedEventArgs e)
        {
            ComponentRenamed?.Invoke(sender, e);
        }

        public event EventHandler<ComponentEventArgs> ComponentActivated;
        private void ComponentActivatedHandler(object sender, ComponentEventArgs e)
        {
            ComponentActivated?.Invoke(sender, e);
        }

        public event EventHandler<ComponentEventArgs> ComponentSelected;
        private void ComponentSelectedHandler(object sender, ComponentEventArgs e)
        {
            ComponentSelected?.Invoke(sender, e);
        }

        public event EventHandler<ComponentEventArgs> ComponentReloaded;
        
        private void ComponentReloadedHandler(object sender, ComponentEventArgs e)
        {
            ComponentReloaded?.Invoke(sender, e);
        }

        public event EventHandler<ReferenceEventArgs> ProjectReferenceAdded;

        private void ProjectReferenceAddedHandler(object sender, ReferenceEventArgs e)
        {
            ProjectReferenceAdded?.Invoke(sender, e);
        }

        public event EventHandler<ReferenceEventArgs> ProjectReferenceRemoved;

        private void ProjectReferenceRemovedHandler(object sender, ReferenceEventArgs e)
        {
            ProjectReferenceRemoved?.Invoke(sender, e);
        }

        public event EventHandler EventsTerminated;

        #region IDisposable

        private bool _disposed;
        /// <remarks>
        /// This is a not a true implementation of IDisposable pattern
        /// because the method is made private and is available only
        /// via the static method <see cref="Terminate"/> to provide
        /// a single point of entry for disposing the singleton class
        /// </remarks>
        private void Dispose(bool disposing)
        {
            if (!_disposed && _projects != null)
            {
                EventsTerminated?.Invoke(this, EventArgs.Empty);

                var projectIds = _components.Keys.ToArray();
                foreach (var projectid in projectIds)
                {
                    UnregisterProjectHandlers(projectid);
                }
                
                _projects.ProjectActivated -= ProjectActivatedHandler;
                _projects.ProjectRenamed -= ProjectRenamedHandler;
                _projects.ProjectRemoved -= ProjectRemovedHandler;
                _projects.ProjectAdded -= ProjectAddedHandler;
                _projects.DetachEvents();
                _projects.Dispose();
                
                _disposed = true;
            }
        }

        ~VBEEvents()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(false);
        }

        // This code added to correctly implement the disposable pattern.
        private void Dispose()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        #endregion
    }
}
