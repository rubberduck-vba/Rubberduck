using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers
{
    public abstract class SafeRedirectedEventedComWrapper<TSource, TEventSource, TEventInterface> : SafeComWrapper<TSource>, ISafeEventedComWrapper
        where TSource : class
        where TEventSource : class
        where TEventInterface : class
    {
        private const int NotAdvising = 0;
        private readonly object _lock = new object();
        private TEventSource _eventSource; // The event source
        private TEventInterface _eventSink; // The event sink
        private IConnectionPoint _icp; // The connection point
        private int _cookie = NotAdvising; // The cookie for the connection

        protected SafeRedirectedEventedComWrapper(TSource target, bool rewrapping = false) 
            : base(target, rewrapping)
        {
        }

        protected override void Dispose(bool disposing)
        {
            DetachEvents();
            base.Dispose(disposing);
        }

        protected void AttachEvents(IEventSource<TEventSource> eventSource, TEventInterface eventSink)
        {
            Debug.Assert(eventSource != null, "Unable to attach events - eventSource is null");
            Debug.Assert(eventSink != null, "Unable to attach events - eventSink is null");
            if (eventSource == null || eventSink == null)
            {
                return;
            }            

            lock (_lock)
            {
                if (IsWrappingNullReference)
                {
                    return;
                }
                
                // Check that events not already attached
                if (_eventSource != null || _eventSink != null)
                {                   
                    return;
                }

                _eventSource = eventSource.EventSource;
                _eventSink = eventSink;

                // Call QueryInterface for IConnectionPointContainer
                if (!(_eventSource is IConnectionPointContainer icpc))
                {
                    Debug.Assert(false, $"Unable to attach events - supplied type {_eventSource.GetType().Name} is not a connection point container.");
                    return;
                }

                // Find the connection point for the source interface
                var g = typeof(TEventInterface).GUID;
                icpc.FindConnectionPoint(ref g, out _icp);
                _icp.Advise(_eventSink, out _cookie);
            }
        }

        public abstract void AttachEvents();

        public void DetachEvents()
        {
            lock (_lock)
            {
                if (_icp != null)
                {
                    if (_cookie != NotAdvising)
                    {
                        _icp.Unadvise(_cookie);
                        _cookie = NotAdvising;
                    }

                    Marshal.ReleaseComObject(_icp);
                    _icp = null;
                }

                if (_eventSource != null)
                {
                    Marshal.ReleaseComObject(_eventSource);
                    _eventSource = null;
                }
            }
        }
    }
}
