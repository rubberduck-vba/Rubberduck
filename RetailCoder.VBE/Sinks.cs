using System;
using System.Collections.Generic;
using System.Runtime.InteropServices.ComTypes;
using System.Threading.Tasks;
using Microsoft.Vbe.Interop;
using Rubberduck.Common.Dispatch;
using Rubberduck.Parsing;
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

        public bool IsEnabled { get; set; }

        public Sinks(VBE vbe)
        {
            IsEnabled = true;

            _sink = new VBProjectsEventsSink();
            var connectionPointContainer = (IConnectionPointContainer)vbe.VBProjects;
            var interfaceId = typeof(_dispVBProjectsEvents).GUID;
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

        private void _sink_ProjectActivated(object sender, DispatcherEventArgs<VBProject> e)
        {
            if (!IsEnabled) { return; }
            var projectId = e.Item.HelpFile;

            Task.Run(() =>
            {
                var handler = ProjectActivated;
                if (handler != null)
                {
                    handler(sender, new ProjectEventArgs(projectId, e.Item));
                }
            });
        }

        private void _sink_ProjectAdded(object sender, DispatcherEventArgs<VBProject> e)
        {
            if (!IsEnabled) { return; }

            e.Item.AssignProjectId();
            var projectId = e.Item.HelpFile;

            RegisterComponentsEventSink(e.Item.VBComponents, projectId);

            Task.Run(() =>
            {
                var handler = ProjectAdded;
                if (handler != null)
                {
                    handler(sender, new ProjectEventArgs(projectId, e.Item));
                }
            });
        }

        private void _sink_ProjectRemoved(object sender, DispatcherEventArgs<VBProject> e)
        {
            UnregisterComponentsEventSink(e.Item.HelpFile);
            if (!IsEnabled) { return; }

            var projectId = e.Item.HelpFile;

            Task.Run(() =>
            {
                var handler = ProjectRemoved;
                if (handler != null)
                {
                    handler(sender, new ProjectEventArgs(projectId, e.Item));
                }
            });
        }

        private void _sink_ProjectRenamed(object sender, DispatcherRenamedEventArgs<VBProject> e)
        {
            if (!IsEnabled) { return; }

            var projectId = e.Item.HelpFile;
            var oldName = e.OldName;

            Task.Run(() =>
            {
                var handler = ProjectRenamed;
                if (handler != null)
                {
                    handler(sender, new ProjectRenamedEventArgs(projectId, e.Item, oldName));
                }
            });
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

            var connectionPointContainer = (IConnectionPointContainer)components;
            var interfaceId = typeof(_dispVBComponentsEvents).GUID;

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

        private void ComponentsSink_ComponentActivated(object sender, DispatcherEventArgs<VBComponent> e)
        {
            if (!IsEnabled) { return; }

            var projectId = e.Item.Collection.Parent.HelpFile;

            var handler = ComponentActivated;
            if (handler != null)
            {
                handler(sender, new ComponentEventArgs(projectId, e.Item.Collection.Parent, e.Item));
            }
        }

        private void ComponentsSink_ComponentAdded(object sender, DispatcherEventArgs<VBComponent> e)
        {
            if (!IsEnabled) { return; }

            var projectId = e.Item.Collection.Parent.HelpFile;
            var componentName = e.Item.Name;

            var handler = ComponentAdded;
            if (handler != null)
            {
                handler(sender, new ComponentEventArgs(projectId, e.Item.Collection.Parent, e.Item));
            }
        }

        private void ComponentsSink_ComponentReloaded(object sender, DispatcherEventArgs<VBComponent> e)
        {
            if (!IsEnabled) { return; }

            var projectId = e.Item.Collection.Parent.HelpFile;
            var componentName = e.Item.Name;

            var handler = ComponentReloaded;
            if (handler != null)
            {
                handler(sender, new ComponentEventArgs(projectId, e.Item.Collection.Parent, e.Item));
            }
        }

        private void ComponentsSink_ComponentRemoved(object sender, DispatcherEventArgs<VBComponent> e)
        {
            if (!IsEnabled) { return; }

            var projectId = e.Item.Collection.Parent.HelpFile;
            var componentName = e.Item.Name;

            var handler = ComponentRemoved;
            if (handler != null)
            {
                handler(sender, new ComponentEventArgs(projectId, e.Item.Collection.Parent, e.Item));
            }
        }

        private void ComponentsSink_ComponentRenamed(object sender, DispatcherRenamedEventArgs<VBComponent> e)
        {
            if (!IsEnabled) { return; }

            var projectId = e.Item.Collection.Parent.HelpFile;
            var componentName = e.Item.Name;
            var oldName = e.OldName;

            var handler = ComponentRenamed;
            if (handler != null)
            {
                handler(sender, new ComponentRenamedEventArgs(projectId, e.Item.Collection.Parent, e.Item, e.OldName));
            }
        }

        private void ComponentsSink_ComponentSelected(object sender, DispatcherEventArgs<VBComponent> e)
        {
            if (!IsEnabled) { return; }

            var projectId = e.Item.Collection.Parent.HelpFile;
            var componentName = e.Item.Name;

            var handler = ComponentSelected;
            if (handler != null)
            {
                handler(sender, new ComponentEventArgs(projectId, e.Item.Collection.Parent, e.Item));
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