using System;
using System.Collections.Generic;
using System.Runtime.InteropServices.ComTypes;
using Rubberduck.Common.Dispatch;
using Rubberduck.Parsing;
using Rubberduck.VBEditor.DisposableWrappers;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck
{
    public class ProjectEventArgs : EventArgs, IProjectEventArgs
    {
        public ProjectEventArgs(string projectId, VBProject project)
        {
            _projectId = projectId;
            _project = project;
        }

        private readonly string _projectId;
        public string ProjectId { get { return _projectId; } }

        private readonly VBProject _project;
        public VBProject Project { get { return _project; } }
    }

    public class ProjectRenamedEventArgs : ProjectEventArgs, IProjectRenamedEventArgs
    {
        public ProjectRenamedEventArgs(string projectId, VBProject project, string oldName) : base(projectId, project)
        {
            _oldName = oldName;
        }

        private readonly string _oldName;
        public string OldName { get { return _oldName; } }
    }

    public class ComponentEventArgs : EventArgs, IComponentEventArgs
    {
        public ComponentEventArgs(string projectId, VBProject project, VBComponent component)
        {
            _projectId = projectId;
            _project = project;
            _component = component;
        }

        private readonly string _projectId;
        public string ProjectId { get { return _projectId; } }

        private readonly VBProject _project;
        public VBProject Project { get { return _project; } }

        private readonly VBComponent _component;
        public VBComponent Component { get { return _component; } }

    }

    public class ComponentRenamedEventArgs : ComponentEventArgs, IComponentRenamedEventArgs
    {
        public ComponentRenamedEventArgs(string projectId, VBProject project, VBComponent component, string oldName)
            : base(projectId, project, component)
        {
            _oldName = oldName;
        }

        private readonly string _oldName;
        public string OldName { get { return _oldName; } }
    }

    public class Sinks : ISinks, IDisposable
    {
        private VBProjectsEventsSink _sink;
        private readonly IConnectionPoint _projectsEventsConnectionPoint;
        private readonly int _projectsEventsCookie;

        private readonly Dictionary<string, VBComponentsEventsSink> _componentsEventsSinks =
            new Dictionary<string, VBComponentsEventsSink>();
        private readonly IDictionary<string, Tuple<IConnectionPoint, int>> _componentsEventsConnectionPoints =
            new Dictionary<string, Tuple<IConnectionPoint, int>>();

        public bool ComponentSinksEnabled { get; set; }

        public Sinks(VBE vbe)
        {
            ComponentSinksEnabled = true;

            _sink = new VBProjectsEventsSink();
            var connectionPointContainer = (IConnectionPointContainer)vbe.VBProjects.ComObject;
            var interfaceId = typeof(Microsoft.Vbe.Interop._dispVBProjectsEvents).GUID;
            connectionPointContainer.FindConnectionPoint(ref interfaceId, out _projectsEventsConnectionPoint);
            _projectsEventsConnectionPoint.Advise(_sink, out _projectsEventsCookie);

            _sink.ProjectActivated += _sink_ProjectActivated;
            _sink.ProjectAdded += _sink_ProjectAdded;
            _sink.ProjectRemoved += _sink_ProjectRemoved;
            _sink.ProjectRenamed += _sink_ProjectRenamed;
        }

        #region ProjectEvents

        public event EventHandler<IProjectEventArgs> ProjectActivated;
        public event EventHandler<IProjectEventArgs> ProjectAdded;
        public event EventHandler<IProjectEventArgs> ProjectRemoved;
        public event EventHandler<IProjectRenamedEventArgs> ProjectRenamed;

        private void _sink_ProjectActivated(object sender, DispatcherEventArgs<Microsoft.Vbe.Interop.VBProject> e)
        {
            var projectId = e.Item.HelpFile;
            
            var handler = ProjectActivated;
            if (handler != null)
            {
                handler(sender, new ProjectEventArgs(projectId, new VBProject(e.Item)));
            }
        }

        private void _sink_ProjectAdded(object sender, DispatcherEventArgs<Microsoft.Vbe.Interop.VBProject> e)
        {
            using (var project = new VBProject(e.Item))
            {
                project.AssignProjectId();
                var projectId = project.HelpFile;

                RegisterComponentsEventSink(project.VBComponents, projectId);

                var handler = ProjectAdded;
                if (handler != null)
                {
                    handler(sender, new ProjectEventArgs(projectId, project));
                }
            }
        }

        private void _sink_ProjectRemoved(object sender, DispatcherEventArgs<Microsoft.Vbe.Interop.VBProject> e)
        {
            using (var project = new VBProject(e.Item))
            {
                var projectId = project.HelpFile;
                UnregisterComponentsEventSink(projectId);

                var handler = ProjectRemoved;
                if (handler != null)
                {
                    handler(sender, new ProjectEventArgs(projectId, project));
                }
            }
        }

        private void _sink_ProjectRenamed(object sender, DispatcherRenamedEventArgs<Microsoft.Vbe.Interop.VBProject> e)
        {
            using (var project = new VBProject(e.Item))
            {
                var projectId = project.HelpFile;
                var oldName = e.OldName;

                var handler = ProjectRenamed;
                if (handler != null)
                {
                    handler(sender, new ProjectRenamedEventArgs(projectId, project, oldName));
                }
            }
        }

        #endregion

        #region ComponentEvents
        private void RegisterComponentsEventSink(VBComponents components, string projectId)
        {
            if (_componentsEventsSinks.ContainsKey(projectId))
            {
                // already registered - this is caused by the initial load+rename of a project in the VBE
                return;
            }

            var connectionPointContainer = (IConnectionPointContainer)components.ComObject;
            var interfaceId = typeof(Microsoft.Vbe.Interop._dispVBComponentsEvents).GUID;

            IConnectionPoint connectionPoint;
            connectionPointContainer.FindConnectionPoint(ref interfaceId, out connectionPoint);

            var componentsSink = new VBComponentsEventsSink();
            componentsSink.ComponentActivated += ComponentsSink_ComponentActivated;
            componentsSink.ComponentAdded += ComponentsSink_ComponentAdded;
            componentsSink.ComponentReloaded += ComponentsSink_ComponentReloaded;
            componentsSink.ComponentRemoved += ComponentsSink_ComponentRemoved;
            componentsSink.ComponentRenamed += ComponentsSink_ComponentRenamed;
            componentsSink.ComponentSelected += ComponentsSink_ComponentSelected;
            _componentsEventsSinks.Add(projectId, componentsSink);

            int cookie;
            connectionPoint.Advise(componentsSink, out cookie);

            _componentsEventsConnectionPoints.Add(projectId, Tuple.Create(connectionPoint, cookie));
        }

        private void UnregisterComponentsEventSink(string projectId)
        {
            var componentEventSink = _componentsEventsSinks[projectId];

            componentEventSink.ComponentActivated -= ComponentsSink_ComponentActivated;
            componentEventSink.ComponentAdded -= ComponentsSink_ComponentAdded;
            componentEventSink.ComponentReloaded -= ComponentsSink_ComponentReloaded;
            componentEventSink.ComponentRemoved -= ComponentsSink_ComponentRemoved;
            componentEventSink.ComponentRenamed -= ComponentsSink_ComponentRenamed;
            componentEventSink.ComponentSelected -= ComponentsSink_ComponentSelected;
            _componentsEventsSinks.Remove(projectId);

            var componentConnectionPoint = _componentsEventsConnectionPoints[projectId];
            componentConnectionPoint.Item1.Unadvise(componentConnectionPoint.Item2);

            _componentsEventsConnectionPoints.Remove(projectId);
        }

        public event EventHandler<IComponentEventArgs> ComponentActivated;
        public event EventHandler<IComponentEventArgs> ComponentAdded;
        public event EventHandler<IComponentEventArgs> ComponentReloaded;
        public event EventHandler<IComponentEventArgs> ComponentRemoved;
        public event EventHandler<IComponentRenamedEventArgs> ComponentRenamed;
        public event EventHandler<IComponentEventArgs> ComponentSelected;

        private void ComponentsSink_ComponentActivated(object sender, DispatcherEventArgs<Microsoft.Vbe.Interop.VBComponent> e)
        {
            if (!ComponentSinksEnabled) { return; }
            using (var component = new VBComponent(e.Item))
            using (var components = component.Collection)
            using (var project = components.Parent)
            {
                var projectId = project.HelpFile;

                var handler = ComponentActivated;
                if (handler != null)
                {
                    handler(sender, new ComponentEventArgs(projectId, project, component));
                }
            }
        }

        private void ComponentsSink_ComponentAdded(object sender, DispatcherEventArgs<Microsoft.Vbe.Interop.VBComponent> e)
        {
            if (!ComponentSinksEnabled) { return; }
            using (var component = new VBComponent(e.Item))
            using (var components = component.Collection)
            using (var project = components.Parent)
            {
                var projectId = project.HelpFile;

                var handler = ComponentAdded;
                if (handler != null)
                {
                    handler(sender, new ComponentEventArgs(projectId, project, component));
                }
            }
        }

        private void ComponentsSink_ComponentReloaded(object sender, DispatcherEventArgs<Microsoft.Vbe.Interop.VBComponent> e)
        {
            if (!ComponentSinksEnabled) { return; }
            using (var component = new VBComponent(e.Item))
            using (var components = component.Collection)
            using (var project = components.Parent)
            {
                var projectId = project.HelpFile;

                var handler = ComponentReloaded;
                if (handler != null)
                {
                    handler(sender, new ComponentEventArgs(projectId, project, component));
                }
            }
        }

        private void ComponentsSink_ComponentRemoved(object sender, DispatcherEventArgs<Microsoft.Vbe.Interop.VBComponent> e)
        {
            if (!ComponentSinksEnabled) { return; }
            using (var component = new VBComponent(e.Item))
            using (var components = component.Collection)
            using (var project = components.Parent)
            {
                var projectId = project.HelpFile;

                var handler = ComponentRemoved;
                if (handler != null)
                {
                    handler(sender, new ComponentEventArgs(projectId, project, component));
                }
            }
        }

        private void ComponentsSink_ComponentRenamed(object sender, DispatcherRenamedEventArgs<Microsoft.Vbe.Interop.VBComponent> e)
        {
            if (!ComponentSinksEnabled) { return; }
            using (var component = new VBComponent(e.Item))
            using (var components = component.Collection)
            using (var project = components.Parent)
            {
                var projectId = project.HelpFile;

                var handler = ComponentRenamed;
                if (handler != null)
                {
                    handler(sender, new ComponentRenamedEventArgs(projectId, project, component, e.OldName));
                }
            }
        }

        private void ComponentsSink_ComponentSelected(object sender, DispatcherEventArgs<Microsoft.Vbe.Interop.VBComponent> e)
        {
            if (!ComponentSinksEnabled) { return; }
            using (var component = new VBComponent(e.Item))
            using (var components = component.Collection)
            using (var project = components.Parent)
            {
                var projectId = project.HelpFile;

                var handler = ComponentSelected;
                if (handler != null)
                {
                    handler(sender, new ComponentEventArgs(projectId, project, component));
                }
            }
        }
        #endregion

        public void Dispose()
        {
            if (_sink != null)
            {
                _sink.ProjectAdded -= _sink_ProjectAdded;
                _sink.ProjectRemoved -= _sink_ProjectRemoved;
                _sink.ProjectActivated -= _sink_ProjectActivated;
                _sink.ProjectRenamed -= _sink_ProjectRenamed;
                _sink = null;
            }

            if (_projectsEventsConnectionPoint != null)
            {
                _projectsEventsConnectionPoint.Unadvise(_projectsEventsCookie);
            }

            foreach (var item in _componentsEventsSinks)
            {
                item.Value.ComponentActivated -= ComponentsSink_ComponentActivated;
                item.Value.ComponentAdded -= ComponentsSink_ComponentAdded;
                item.Value.ComponentReloaded -= ComponentsSink_ComponentReloaded;
                item.Value.ComponentRemoved -= ComponentsSink_ComponentRemoved;
                item.Value.ComponentRenamed -= ComponentsSink_ComponentRenamed;
                item.Value.ComponentSelected -= ComponentsSink_ComponentSelected;
            }

            foreach (var item in _componentsEventsConnectionPoints)
            {
                item.Value.Item1.Unadvise(item.Value.Item2);
            }
        }
    }
}