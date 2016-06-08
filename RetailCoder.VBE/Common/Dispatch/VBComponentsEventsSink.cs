using System;
using Microsoft.Vbe.Interop;

namespace Rubberduck.Common.Dispatch
{
    public class VBComponentsEventsSink : _dispVBComponentsEvents
    {
        public event EventHandler<DispatcherEventArgs<VBComponent>> ComponentAdded;
        public void ItemAdded(VBComponent VBComponent)
        {
            OnDispatch(ComponentAdded, VBComponent);
        }

        public event EventHandler<DispatcherEventArgs<VBComponent>> ComponentRemoved;
        public void ItemRemoved(VBComponent VBComponent)
        {
            OnDispatch(ComponentRemoved, VBComponent);
        }

        public event EventHandler<DispatcherRenamedEventArgs<VBComponent>> ComponentRenamed;
        public void ItemRenamed(VBComponent VBComponent, string OldName)
        {
            var handler = ComponentRenamed;
            if (handler != null)
            {
                handler.Invoke(this, new DispatcherRenamedEventArgs<VBComponent>(VBComponent, OldName));
            }
        }

        public event EventHandler<DispatcherEventArgs<VBComponent>> ComponentSelected;
        public void ItemSelected(VBComponent VBComponent)
        {
            OnDispatch(ComponentSelected, VBComponent);
        }

        public event EventHandler<DispatcherEventArgs<VBComponent>> ComponentActivated;
        public void ItemActivated(VBComponent VBComponent)
        {
            OnDispatch(ComponentActivated, VBComponent);
        }

        public event EventHandler<DispatcherEventArgs<VBComponent>> ComponentReloaded;
        public void ItemReloaded(VBComponent VBComponent)
        {
            OnDispatch(ComponentReloaded, VBComponent);
        }

        private void OnDispatch(EventHandler<DispatcherEventArgs<VBComponent>> dispatched, VBComponent component)
        {
            var handler = dispatched;
            if (handler != null)
            {
                handler.Invoke(this, new DispatcherEventArgs<VBComponent>(component));
            }
        }
    }
}
