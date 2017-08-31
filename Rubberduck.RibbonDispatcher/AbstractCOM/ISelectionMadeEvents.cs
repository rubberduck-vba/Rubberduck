using System;
using System.Runtime.InteropServices;

namespace Rubberduck.RibbonDispatcher.AbstractCOM {
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    [Guid(RubberduckGuid.ISelectedEvents)]
    public interface ISelectionMadeEvents {
        /// <summary>TODO</summary>
        [DispId(1)]
        void SelectionMade(string itemId, int itemIndex);
    }
}
