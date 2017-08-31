using System;
using System.Runtime.InteropServices;

namespace Rubberduck.RibbonDispatcher.AbstractCOM {
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    [Guid(RubberduckGuid.IClickedEvents)]
    public interface IClickedEvents {
        /// <summary>TODO</summary>
        [DispId(DispIds.Clicked)]
        void Clicked();
    }
}
