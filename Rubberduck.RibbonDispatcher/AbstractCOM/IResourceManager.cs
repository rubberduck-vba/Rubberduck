using System;
using System.Runtime.InteropServices;

namespace Rubberduck.RibbonDispatcher.AbstractCOM {
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(RubberduckGuid.IResourceManager)]
    public interface IResourceManager {
        /// <summary>TODO</summary>
        [DispId(1)]
        string GetCurrentUIString(string name);
    }
}
