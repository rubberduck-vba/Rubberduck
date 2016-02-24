using System;
using Microsoft.Vbe.Interop;

namespace Rubberduck.Common.Dispatch
{
    public class ReferencesEventsSink : _dispReferencesEvents
    {
        public event EventHandler<DispatcherEventArgs<Reference>> ReferenceAdded;
        public void ItemAdded(Reference Reference)
        {
            OnDispatch(ReferenceAdded, Reference);
        }

        public event EventHandler<DispatcherEventArgs<Reference>> ReferenceRemoved;
        public void ItemRemoved(Reference Reference)
        {
            OnDispatch(ReferenceRemoved, Reference);
        }

        private void OnDispatch(EventHandler<DispatcherEventArgs<Reference>> dispatched, Reference reference)
        {
            var handler = dispatched;
            if (handler != null)
            {
                handler.Invoke(this, new DispatcherEventArgs<Reference>(reference));
            }
        }
    }
}