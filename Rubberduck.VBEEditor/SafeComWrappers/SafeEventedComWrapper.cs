using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace Rubberduck.VBEditor.SafeComWrappers
{
    public abstract class SafeEventedComWrapper<TSource, TEventInterface> : SafeComWrapper<TSource>
        where TSource : class
        where TEventInterface : class
    {
        private IConnectionPoint _icp; // The connection point
        private int _cookie = -1;     // The cookie for the connection

        protected SafeEventedComWrapper(TSource target, bool rewrapping = false) : base(target, rewrapping)
        {
            AttachEvents();
        }

        protected override void Dispose(bool disposing)
        {
            DetatchEvents();
            base.Dispose(disposing);
        }

        private void AttachEvents()
        {
            // Call QueryInterface for IConnectionPointContainer
            var icpc = (IConnectionPointContainer)Target;

            // Find the connection point for the source interface
            var g = typeof(TEventInterface).GUID;
            icpc.FindConnectionPoint(ref g, out _icp);

            // Pass a pointer to the host to the connection point
            _icp.Advise(this as TEventInterface, out _cookie);
        }

        private void DetatchEvents()
        {
            if (_cookie != -1)
            {
                _icp.Unadvise(_cookie);
            }
            if (_icp != null)
            {
                Marshal.ReleaseComObject(_icp);
            }
        }
    }
}
