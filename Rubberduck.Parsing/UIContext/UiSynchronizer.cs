using System;
using System.Runtime.InteropServices;
using System.Threading;

namespace Rubberduck.Parsing.UIContext
{
    public static class UiSynchronizer
    {
        private enum DllVersion
        {
            Unknown,
            Vbe6,
            Vbe7
        }

        private static DllVersion _version;

        private static SynchronizationContext UiContext { get; set; }

        static UiSynchronizer()
        {
            _version = DllVersion.Unknown;
        }

        public static void Initialize()
        {
            if (UiContext == null)
            {
                UiContext = SynchronizationContext.Current;
            }
        }

        /// <summary>
        /// Used to pump any pending COM messages. This should be used only as a part of
        /// synchronizing or to effect a block until all other threads has finished with 
        /// their pending COM calls. This should be used by the UI thread **ONLY**; 
        /// otherwise execptions will be thrown.
        /// </summary>
        /// <remarks>
        /// Typical use would be within a event handler for an event belonging to a COM 
        /// object which require some synchronization with COM accesses from other threads. 
        /// Events raised by COM are on UI thread by definition so the call stack originating
        /// from COM objects' events can use this method.
        /// </remarks>
        /// <returns>Count of open forms which is always zero for VBA hosts but may be nonzero for VB6 projects.</returns>
        public static int DoEvents()
        {
            CheckContext();
            
            return ExecuteDoEvents();
        }

        private static int ExecuteDoEvents()
        {
            switch (_version)
            {
                case DllVersion.Vbe7:
                    return rtcDoEvents7();
                case DllVersion.Vbe6:
                    return rtcDoEvents6();
                default:
                    return DetermineVersionAndExecute();
            }
        }

        private static int DetermineVersionAndExecute()
        {
            int result;
            try

            {
                result = rtcDoEvents7();
                _version = DllVersion.Vbe7;
            }
            catch
            {
                try
                {
                    result = rtcDoEvents6();
                    _version = DllVersion.Vbe6;
                }
                catch
                {
                    // we shouldn't be here.... Rubberduck is a VBA add-in, so how the heck could it have loaded without a VBE dll?!?
                    throw new InvalidOperationException("Cannot execute DoEvents; the VBE dll could not be located.");
                }
            }

            return result;
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

        [DllImport("vbe6.dll", EntryPoint = "rtcDoEvents")]
        private static extern int rtcDoEvents6();

        [DllImport("vbe7.dll", EntryPoint = "rtcDoEvents")]
        private static extern int rtcDoEvents7();
    }
}
