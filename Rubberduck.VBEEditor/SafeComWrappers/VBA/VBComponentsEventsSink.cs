using System;
using System.Diagnostics.CodeAnalysis;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    [SuppressMessage("ReSharper", "InconsistentNaming")]
    public class VBComponentsEventsSink : VB._dispVBComponentsEvents, IVBComponentsEventsSink
    {
        public event EventHandler<DispatcherEventArgs<IVBComponent>> ComponentAdded;
        public void ItemAdded(VB.VBComponent VBComponent)
        {
            OnDispatch(ComponentAdded, VBComponent);
        }

        public event EventHandler<DispatcherEventArgs<IVBComponent>> ComponentRemoved;
        public void ItemRemoved(VB.VBComponent VBComponent)
        {
            OnDispatch(ComponentRemoved, VBComponent);
        }

        public event EventHandler<DispatcherRenamedEventArgs<IVBComponent>> ComponentRenamed;
        public void ItemRenamed(VB.VBComponent VBComponent, string OldName)
        {
            var handler = ComponentRenamed;
            if (handler != null)
            {
                handler.Invoke(this, new DispatcherRenamedEventArgs<IVBComponent>(new VBComponent(VBComponent), OldName));
            }
        }

        public event EventHandler<DispatcherEventArgs<IVBComponent>> ComponentSelected;
        public void ItemSelected(VB.VBComponent VBComponent)
        {
            OnDispatch(ComponentSelected, VBComponent);
        }

        public event EventHandler<DispatcherEventArgs<IVBComponent>> ComponentActivated;
        public void ItemActivated(VB.VBComponent VBComponent)
        {
            OnDispatch(ComponentActivated, VBComponent);
        }

        public event EventHandler<DispatcherEventArgs<IVBComponent>> ComponentReloaded;
        public void ItemReloaded(VB.VBComponent VBComponent)
        {
            OnDispatch(ComponentReloaded, VBComponent);
        }

        private void OnDispatch(EventHandler<DispatcherEventArgs<IVBComponent>> dispatched, VB.VBComponent component)
        {
            var handler = dispatched;
            if (handler != null)
            {
                handler.Invoke(this, new DispatcherEventArgs<IVBComponent>(new VBComponent(component)));
            }
        }
    }
}
