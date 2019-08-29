using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.Events
{
    public sealed class VbeEvents : IVbeEvents
    {
        private static VbeEvents _instance;
        private static readonly object _instanceLock = new object();
        private readonly IVBProjects _projects;
        private readonly Dictionary<string, IVBComponents> _components;
        private readonly Dictionary<string, IReferences> _references;

        private static long _terminated;
        private const long _true = 1;
        private const long _false = 0;

        public static VbeEvents Initialize(IVBE vbe)
        {
            lock (_instanceLock)
            {
                if (_instance == null)
                {
                    _instance = new VbeEvents(vbe);
                    Interlocked.Exchange(ref _terminated, _false);
                }
            }

            return _instance;
        }

        public static void Terminate()
        {
            lock (_instanceLock)
            {
                if (_instance == null)
                {
                    return;
                }

                _instance.TerminateInternal();
                _instance = null;
            }
        }

        private VbeEvents(IVBE vbe)
        {
            _components = new Dictionary<string, IVBComponents>();
            _references = new Dictionary<string, IReferences>();

            if (_projects != null)
            {
                return;
            }
            
            _projects = vbe.VBProjects;

            if (_projects.IsWrappingNullReference)
            {
                return;
            }

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

        public bool Terminated => Interlocked.Read(ref _terminated) == _true;

        private void TerminateInternal()
        {
            // If we fail, we at least should advertise that we're now dead
            Interlocked.Exchange(ref _terminated, _true);

            EventsTerminated?.Invoke(this, EventArgs.Empty);
            EventsTerminated = delegate { };
            var projectIds = _components.Keys.ToArray();
            foreach (var projectid in projectIds)
            {
                UnregisterProjectHandlers(projectid);
            }

            if (_projects.IsWrappingNullReference)
            {
                return;
            }

            _projects.ProjectActivated -= ProjectActivatedHandler;
            _projects.ProjectRenamed -= ProjectRenamedHandler;
            _projects.ProjectRemoved -= ProjectRemovedHandler;
            _projects.ProjectAdded -= ProjectAddedHandler;
            _projects.DetachEvents();
            _projects.Dispose();
        }
    }
}
