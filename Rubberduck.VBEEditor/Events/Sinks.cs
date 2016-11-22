using System;
using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.Events
{
    public class Sinks : ISinks, IDisposable
    {
        private readonly IVBE _vbe;

        private readonly ConnectionPointInfo _projectEventsInfo;
        private IVBProjectsEventsSink _projectsSink;

        private readonly IDictionary<string, IVBComponentsEventsSink> _componentsEventsSinks = new Dictionary<string, IVBComponentsEventsSink>();
        private readonly IDictionary<string, ConnectionPointInfo> _componentEventsInfo = new Dictionary<string, ConnectionPointInfo>();

        public bool ComponentSinksEnabled { get; set; }

        public Sinks(IVBE vbe)
        {
            _vbe = vbe;
            _projectsSink = _vbe.VBProjects.Events;
            _projectEventsInfo = new ConnectionPointInfo(_vbe.VBProjects.ConnectionPoint);
            ComponentSinksEnabled = true;
        }

        public void Start()
        {
            if (!_projectEventsInfo.HasConnectionPoint)  { return; }

            _projectsSink.ProjectActivated += _sink_ProjectActivated;
            _projectsSink.ProjectAdded += _sink_ProjectAdded;
            _projectsSink.ProjectRemoved += _sink_ProjectRemoved;
            _projectsSink.ProjectRenamed += _sink_ProjectRenamed;
            _projectEventsInfo.Advise(_projectsSink);
            foreach (var project in _vbe.VBProjects)
            {
                RegisterComponentsEventSink(project.VBComponents, project.ProjectId);
            }
        }

        public void Stop()
        {
            _projectsSink.ProjectActivated -= _sink_ProjectActivated;
            _projectsSink.ProjectAdded -= _sink_ProjectAdded;
            _projectsSink.ProjectRemoved -= _sink_ProjectRemoved;
            _projectsSink.ProjectRenamed -= _sink_ProjectRenamed;
            _projectEventsInfo.Unadvise();
            
            foreach (var project in _vbe.VBProjects)
            {
                UnregisterComponentsEventSink(project.ProjectId);
            }
        }

        private void RegisterComponentsEventSink(IVBComponents components, string projectId)
        {
            if (projectId == null || _componentsEventsSinks.ContainsKey(projectId))
            {
                // already registered - this is caused by the initial load+rename of a project in the VBE
                return;
            }

            var connectionPoint = components.ConnectionPoint;
            var componentsSink = components.Events;
            componentsSink.ComponentActivated += ComponentsSink_ComponentActivated;
            componentsSink.ComponentAdded += ComponentsSink_ComponentAdded;
            componentsSink.ComponentReloaded += ComponentsSink_ComponentReloaded;
            componentsSink.ComponentRemoved += ComponentsSink_ComponentRemoved;
            componentsSink.ComponentRenamed += ComponentsSink_ComponentRenamed;
            componentsSink.ComponentSelected += ComponentsSink_ComponentSelected;
            var info = new ConnectionPointInfo(connectionPoint);
            info.Advise(componentsSink);
            _componentEventsInfo.Add(projectId, info);
        }

        private void UnregisterComponentsEventSink(string projectId)
        {
            if (projectId == null || !_componentsEventsSinks.ContainsKey(projectId)) { return; }

            var componentEventSink = _componentsEventsSinks[projectId];
            var info = _componentEventsInfo[projectId];
            info.Unadvise();
            componentEventSink.ComponentActivated -= ComponentsSink_ComponentActivated;
            componentEventSink.ComponentAdded -= ComponentsSink_ComponentAdded;
            componentEventSink.ComponentReloaded -= ComponentsSink_ComponentReloaded;
            componentEventSink.ComponentRemoved -= ComponentsSink_ComponentRemoved;
            componentEventSink.ComponentRenamed -= ComponentsSink_ComponentRenamed;
            componentEventSink.ComponentSelected -= ComponentsSink_ComponentSelected;
            _componentsEventsSinks.Remove(projectId);
            _componentEventsInfo.Remove(projectId);
        }

        #region ProjectEvents

        public event EventHandler<ProjectEventArgs> ProjectActivated;
        public event EventHandler<ProjectEventArgs> ProjectAdded;
        public event EventHandler<ProjectEventArgs> ProjectRemoved;
        public event EventHandler<ProjectRenamedEventArgs> ProjectRenamed;

        private void _sink_ProjectActivated(object sender, DispatcherEventArgs<IVBProject> e)
        {
            if (!_vbe.IsInDesignMode)  { return; }

            var project = e.Item;
            var projectId = project.ProjectId;
            
            var handler = ProjectActivated;
            if (handler != null)
            {
                handler(sender, new ProjectEventArgs(projectId, project));
            }
        }

        private void _sink_ProjectAdded(object sender, DispatcherEventArgs<IVBProject> e)
        {
            if (!_vbe.IsInDesignMode) { return; }

            var project = e.Item;
            project.AssignProjectId();
            var projectId = project.ProjectId;

            RegisterComponentsEventSink(project.VBComponents, projectId);

            var handler = ProjectAdded;
            if (handler != null)
            {
                handler(sender, new ProjectEventArgs(projectId, project));
            }
        }

        private void _sink_ProjectRemoved(object sender, DispatcherEventArgs<IVBProject> e)
        {
            if (!_vbe.IsInDesignMode) { return; }

            var project = e.Item;
            var projectId = project.ProjectId;
            UnregisterComponentsEventSink(projectId);

            var handler = ProjectRemoved;
            if (handler != null)
            {
                handler(sender, new ProjectEventArgs(projectId, project));
            }

        }

        private void _sink_ProjectRenamed(object sender, DispatcherRenamedEventArgs<IVBProject> e)
        {
            if (!_vbe.IsInDesignMode) { return; }

            var project = e.Item;
            var projectId = project.ProjectId;
            var oldName = e.OldName;

            var handler = ProjectRenamed;
            if (handler != null)
            {
                handler(sender, new ProjectRenamedEventArgs(projectId, project, oldName));
            }

        }

        #endregion

        #region ComponentEvents
        public event EventHandler<ComponentEventArgs> ComponentSelected;
        public event EventHandler<ComponentEventArgs> ComponentActivated;
        public event EventHandler<ComponentEventArgs> ComponentAdded;
        public event EventHandler<ComponentEventArgs> ComponentReloaded;
        public event EventHandler<ComponentEventArgs> ComponentRemoved;
        public event EventHandler<ComponentRenamedEventArgs> ComponentRenamed;

        private void ComponentsSink_ComponentActivated(object sender, DispatcherEventArgs<IVBComponent> e)
        {
            if (!ComponentSinksEnabled || !_vbe.IsInDesignMode) {  return; }

            var component = e.Item;
            var components = component.Collection;
            var project = components.Parent;

            var projectId = project.ProjectId;

            var handler = ComponentActivated;
            if (handler != null)
            {
                handler(sender, new ComponentEventArgs(projectId, project, component));
            }

        }

        private void ComponentsSink_ComponentAdded(object sender, DispatcherEventArgs<IVBComponent> e)
        {
            if (!ComponentSinksEnabled || !_vbe.IsInDesignMode) { return; }

            var component = e.Item;
            var components = component.Collection;
            var project = components.Parent;

            var projectId = project.ProjectId;

            var handler = ComponentAdded;
            if (handler != null)
            {
                handler(sender, new ComponentEventArgs(projectId, project, component));
            }

        }

        private void ComponentsSink_ComponentReloaded(object sender, DispatcherEventArgs<IVBComponent> e)
        {
            if (!ComponentSinksEnabled || !_vbe.IsInDesignMode) { return; }

            var component = e.Item;
            var components = component.Collection;
            var project = components.Parent;

            var projectId = project.ProjectId;

            var handler = ComponentReloaded;
            if (handler != null)
            {
                handler(sender, new ComponentEventArgs(projectId, project, component));
            }

        }

        private void ComponentsSink_ComponentRemoved(object sender, DispatcherEventArgs<IVBComponent> e)
        {
            if (!ComponentSinksEnabled || !_vbe.IsInDesignMode) { return; }

            var component = e.Item;
            var components = component.Collection;
            var project = components.Parent;
            var projectId = project.ProjectId;

            var handler = ComponentRemoved;
            if (handler != null)
            {
                handler(sender, new ComponentEventArgs(projectId, project, component));
            }

        }

        private void ComponentsSink_ComponentRenamed(object sender, DispatcherRenamedEventArgs<IVBComponent> e)
        {
            if (!ComponentSinksEnabled || !_vbe.IsInDesignMode) { return; }

            var component = e.Item;
            var components = component.Collection;
            var project = components.Parent;
            var projectId = project.ProjectId;

            var handler = ComponentRenamed;
            if (handler != null)
            {
                handler(sender, new ComponentRenamedEventArgs(projectId, project, component, e.OldName));
            }

        }

        private void ComponentsSink_ComponentSelected(object sender, DispatcherEventArgs<IVBComponent> e)
        {
            if (!ComponentSinksEnabled || !_vbe.IsInDesignMode) { return; }

            var component = e.Item;
            var components = component.Collection;
            var project = components.Parent;

            var projectId = project.ProjectId;

            var handler = ComponentSelected;
            if (handler != null)
            {
                handler(sender, new ComponentEventArgs(projectId, project, component));
            }

        }

        #endregion

        public void Dispose()
        {
            if (_projectsSink != null)
            {
                Stop();
                _projectsSink = null;
            }
        }
    }
}