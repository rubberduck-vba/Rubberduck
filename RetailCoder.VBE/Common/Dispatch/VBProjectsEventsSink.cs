using System;
using Microsoft.Vbe.Interop;

namespace Rubberduck.Common.Dispatch
{
    public class VBProjectsEventsSink : _dispVBProjectsEvents
    {
        public event EventHandler<DispatcherEventArgs<VBProject>> ProjectAdded;
        public void ItemAdded(VBProject VBProject)
        {
            OnDispatch(ProjectAdded, VBProject);
        }

        public event EventHandler<DispatcherEventArgs<VBProject>> ProjectRemoved;
        public void ItemRemoved(VBProject VBProject)
        {
            OnDispatch(ProjectRemoved, VBProject);
        }

        public event EventHandler<DispatcherRenamedEventArgs<VBProject>> ProjectRenamed;
        public void ItemRenamed(VBProject VBProject, string OldName)
        {
            var handler = ProjectRenamed;
            if (handler != null)
            {
                handler.Invoke(this, new DispatcherRenamedEventArgs<VBProject>(VBProject, OldName));
            }
        }

        public event EventHandler<DispatcherEventArgs<VBProject>> ProjectActivated;
        public void ItemActivated(VBProject VBProject)
        {
            OnDispatch(ProjectActivated, VBProject);
        }

        private void OnDispatch(EventHandler<DispatcherEventArgs<VBProject>> dispatched, VBProject project)
        {
            var handler = dispatched;
            if (handler != null)
            {
                handler.Invoke(this, new DispatcherEventArgs<VBProject>(project));
            }
        }
    }
}
