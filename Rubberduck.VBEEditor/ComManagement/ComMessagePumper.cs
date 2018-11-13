using System;
using Rubberduck.VBEditor.Utility;
using Rubberduck.VBEditor.VbeRuntime;

namespace Rubberduck.VBEditor.ComManagement
{
    public interface IComMessagePumper
    {
        int PumpMessages();
    }

    public class ComMessagePumper : IComMessagePumper
    {
        private readonly IVbeNativeApi _runtime;
        private readonly IUiContextProvider _uiContext;

        public ComMessagePumper(IUiContextProvider uiContext, IVbeNativeApi runtime)
        {
            _uiContext = uiContext;
            _runtime = runtime;
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
        public int PumpMessages()
        {
            CheckContext();

            return _runtime.DoEvents();
        }

        private void CheckContext()
        {
            if (!_uiContext.IsExecutingInUiContext())
            {
                throw new InvalidOperationException("ComMessagePumper cannot be used in other threads. Only the UI thread can call methods on the ComMessagePumper");
            }
        }
    }
}
