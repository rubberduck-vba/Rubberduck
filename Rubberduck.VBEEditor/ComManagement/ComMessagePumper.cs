using System;
using System.Threading;
using Rubberduck.VBEditor.ComManagement.VBERuntime;

namespace Rubberduck.VBEditor.ComManagement
{
    public static class ComMessagePumper
    {
        private static SynchronizationContext UiContext { get; set; }
        private static readonly IVBERuntime _runtime;

        public static void Initialize()
        {
            if (UiContext == null)
            {
                UiContext = SynchronizationContext.Current;
            }

            _runtime = new VBERuntimeAccessor();
        }

        /// <summary>
        /// Used to pump any pending COM messages. This should be used only as a part of
        /// synchronizing or to effect a block until all other threads has finished with 
        /// their pending COM calls. This should be used by the UI thread **ONLY**; 
        /// otherwise execptions will be thrown. This is mandatory when the COM objects are
        /// STA which would otherwise deadlock in other threads. 
        /// </summary>
        /// <remarks>
        /// Typical use would be within a event handler for an event belonging to a COM 
        /// object which require some synchronization with COM accesses from other threads. 
        /// Events raised by COM are on UI thread by definition so the call stack originating
        /// from COM objects' events can use this method.
        /// </remarks>
        /// <returns>Count of open forms which is always zero for VBA hosts but may be nonzero for VB6 projects.</returns>
        public static int PumpMessages()
        {
            CheckContext();

            return _runtime.DoEvents();
        }

        private static void CheckContext()
        {
            if (UiContext == null)
            {
                throw new InvalidOperationException("UiSynchronizer is not initialized. Invoke Initialize() from UI thread first.");
            }

            if (UiContext != SynchronizationContext.Current)
            {
                throw new InvalidOperationException("UiSynchronizer cannot be used in other threads. Only the UI thread can call methods on the UiSynchronizer");
            }
        }
    }
}
