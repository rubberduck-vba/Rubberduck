using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers
{
    public abstract class SafeEventedComWrapper<TSource, TEventInterface> : SafeComWrapper<TSource>, ISafeEventedComWrapper
        where TSource : class
        where TEventInterface : class
    {
        private const int NotAdvising = 0;
        private readonly object _lock = new object();
        private IConnectionPoint _icp; // The connection point
        private int _cookie = NotAdvising;     // The cookie for the connection

        protected SafeEventedComWrapper(TSource target, bool rewrapping = false) : base(target, rewrapping)
        {
        }

        protected override void Dispose(bool disposing)
        {
            DetachEvents();
            base.Dispose(disposing);
        }

        public void AttachEvents()
        {
            lock (_lock)
            {
                if (IsWrappingNullReference)
                {
                    return;
                }

                if (_cookie != NotAdvising)
                {
                    return;
                }

                // Call QueryInterface for IConnectionPointContainer
                var icpc = (IConnectionPointContainer) Target;

                // Find the connection point for the source interface
                var g = typeof(TEventInterface).GUID;
                icpc.FindConnectionPoint(ref g, out _icp);

                var sink = this as TEventInterface;

                if (sink == null)
                {
                    throw new InvalidOperationException($"The class {this.GetType()} does not implement the required event interface {typeof(TEventInterface)}");
                }
                
                _icp.Advise(sink, out _cookie);
            }
        }

        public void DetachEvents()
        {
            lock (_lock)
            {
                if (_icp == null)
                {
                    return;
                }

                if (_cookie != NotAdvising)
                {
                    _icp.Unadvise(_cookie);
                    _cookie = NotAdvising;
                }

                Marshal.ReleaseComObject(_icp);
                _icp = null;
            }
        }
    }
}
