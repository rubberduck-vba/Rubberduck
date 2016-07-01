using System;
using System.Collections.Generic;
using System.Runtime.InteropServices.ComTypes;
using Microsoft.Vbe.Interop;
using Rubberduck.Common.Dispatch;
using Rubberduck.Parsing;

namespace Rubberduck
{
    public class Sinks : ISinks, IDisposable
    {
        private VBProjectsEventsSink _sink;
        private readonly IConnectionPoint _projectsEventsConnectionPoint;
        private readonly int _projectsEventsCookie;

        private readonly Dictionary<string, VBComponentsEventsSink> _componentsEventsSinks =
            new Dictionary<string, VBComponentsEventsSink>();
        private readonly IDictionary<string, Tuple<IConnectionPoint, int>> _componentsEventsConnectionPoints =
            new Dictionary<string, Tuple<IConnectionPoint, int>>();

        private readonly IDictionary<string, ReferencesEventsSink> _referencesEventsSinks =
            new Dictionary<string, ReferencesEventsSink>();
        private readonly IDictionary<string, Tuple<IConnectionPoint, int>> _referencesEventsConnectionPoints =
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

        public event EventHandler<IDispatcherEventArgs<VBProject>> ProjectActivated;
        public event EventHandler<IDispatcherEventArgs<VBProject>> ProjectAdded;
        public event EventHandler<IDispatcherEventArgs<VBProject>> ProjectRemoved;
        public event EventHandler<IDispatcherRenamedEventArgs<VBProject>> ProjectRenamed;

        private void _sink_ProjectActivated(object sender, DispatcherEventArgs<VBProject> e)
        {
            if (!IsEnabled) { return; }

            var handler = ProjectActivated;
            if (handler != null)
            {
                handler(sender, e);
            }
        }

        private void _sink_ProjectAdded(object sender, DispatcherEventArgs<VBProject> e)
        {
            if (!IsEnabled) { return; }

            var handler = ProjectAdded;
            if (handler != null)
            {
                handler(sender, e);
            }

            RegisterComponentsEventSink(e.Item.VBComponents, e.Item.HelpFile);
            //RegisterReferencesEventSink(e.Item.References, e.Item.HelpFile);
        }

        private void _sink_ProjectRemoved(object sender, DispatcherEventArgs<VBProject> e)
        {
            if (!IsEnabled) { return; }

            var handler = ProjectRemoved;
            if (handler != null)
            {
                handler(sender, e);
            }

            UnregisterComponentsEventSink(e.Item.HelpFile);
            //UnregisterReferencesEventSink(e.Item.HelpFile);
        }

        private void _sink_ProjectRenamed(object sender, DispatcherRenamedEventArgs<VBProject> e)
        {
            if (!IsEnabled) { return; }

            var handler = ProjectRenamed;
            if (handler != null)
            {
                handler(sender, e);
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

        public event EventHandler<IDispatcherEventArgs<VBComponent>> ComponentActivated;
        public event EventHandler<IDispatcherEventArgs<VBComponent>> ComponentAdded;
        public event EventHandler<IDispatcherEventArgs<VBComponent>> ComponentReloaded;
        public event EventHandler<IDispatcherEventArgs<VBComponent>> ComponentRemoved;
        public event EventHandler<IDispatcherRenamedEventArgs<VBComponent>> ComponentRenamed;
        public event EventHandler<IDispatcherEventArgs<VBComponent>> ComponentSelected;

        private void ComponentsSink_ComponentActivated(object sender, DispatcherEventArgs<VBComponent> e)
        {
            if (!IsEnabled) { return; }

            var handler = ComponentActivated;
            if (handler != null)
            {
                handler(sender, e);
            }
        }

        private void ComponentsSink_ComponentAdded(object sender, DispatcherEventArgs<VBComponent> e)
        {
            if (!IsEnabled) { return; }

            var handler = ComponentAdded;
            if (handler != null)
            {
                handler(sender, e);
            }
        }

        private void ComponentsSink_ComponentReloaded(object sender, DispatcherEventArgs<VBComponent> e)
        {
            if (!IsEnabled) { return; }

            var handler = ComponentReloaded;
            if (handler != null)
            {
                handler(sender, e);
            }
        }

        private void ComponentsSink_ComponentRemoved(object sender, DispatcherEventArgs<VBComponent> e)
        {
            if (!IsEnabled) { return; }

            var handler = ComponentRemoved;
            if (handler != null)
            {
                handler(sender, e);
            }
        }

        private void ComponentsSink_ComponentRenamed(object sender, DispatcherRenamedEventArgs<VBComponent> e)
        {
            if (!IsEnabled) { return; }

            var handler = ComponentRenamed;
            if (handler != null)
            {
                handler(sender, e);
            }
        }

        private void ComponentsSink_ComponentSelected(object sender, DispatcherEventArgs<VBComponent> e)
        {
            if (!IsEnabled) { return; }

            var handler = ComponentSelected;
            if (handler != null)
            {
                handler(sender, e);
            }
        }
        #endregion

        // todo figure out why that cast fails
        //#region ReferenceEvents
        //private void RegisterReferencesEventSink(References references, string projectId)
        //{
        //    if (_referencesEventsSinks.ContainsKey(projectId))
        //    {
        //        // already registered - this is caused by the initial load+rename of a project in the VBE
        //        return;
        //    }

        //    var connectionPointContainer = (IConnectionPointContainer)references;
        //    var interfaceId = typeof(_dispReferencesEvents).GUID;

        //    IConnectionPoint connectionPoint;
        //    connectionPointContainer.FindConnectionPoint(ref interfaceId, out connectionPoint);

        //    var referencesSink = new ReferencesEventsSink();
        //    referencesSink.ReferenceAdded += ReferencesSink_ReferenceAdded;
        //    referencesSink.ReferenceRemoved += ReferencesSink_ReferenceRemoved;
        //    _referencesEventsSinks.Add(projectId, referencesSink);

        //    int cookie;
        //    connectionPoint.Advise(referencesSink, out cookie);

        //    _referencesEventsConnectionPoints.Add(projectId, Tuple.Create(connectionPoint, cookie));
        //}

        //private void UnregisterReferencesEventSink(string projectId)
        //{
        //    var referencesEventSink = _referencesEventsSinks[projectId];

        //    referencesEventSink.ReferenceAdded -= ReferencesSink_ReferenceAdded;
        //    referencesEventSink.ReferenceRemoved -= ReferencesSink_ReferenceRemoved;
        //    _referencesEventsSinks.Remove(projectId);

        //    var referenceConnectionPoint = _referencesEventsConnectionPoints[projectId];
        //    referenceConnectionPoint.Item1.Unadvise(referenceConnectionPoint.Item2);

        //    _referencesEventsConnectionPoints.Remove(projectId);
        //}

        //public event EventHandler<IDispatcherEventArgs<Reference>> ReferenceAdded;
        //public event EventHandler<IDispatcherEventArgs<Reference>> ReferenceRemoved;

        //private void ReferencesSink_ReferenceAdded(object sender, DispatcherEventArgs<Reference> e)
        //{
        //    if (!IsEnabled) { return; }

        //    var handler = ReferenceAdded;
        //    if (handler != null)
        //    {
        //        handler(sender, e);
        //    }
        //}

        //private void ReferencesSink_ReferenceRemoved(object sender, DispatcherEventArgs<Reference> e)
        //{
        //    if (!IsEnabled) { return; }

        //    var handler = ReferenceRemoved;
        //    if (handler != null)
        //    {
        //        handler(sender, e);
        //    }
        //}
        //#endregion

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

            /*foreach (var item in _referencesEventsSinks)
            {
                item.Value.ReferenceAdded -= ReferencesSink_ReferenceAdded;
                item.Value.ReferenceRemoved -= ReferencesSink_ReferenceRemoved;
            }*/

            foreach (var item in _componentsEventsConnectionPoints)
            {
                item.Value.Item1.Unadvise(item.Value.Item2);
            }

            foreach (var item in _referencesEventsConnectionPoints)
            {
                item.Value.Item1.Unadvise(item.Value.Item2);
            }
        }
    }
}