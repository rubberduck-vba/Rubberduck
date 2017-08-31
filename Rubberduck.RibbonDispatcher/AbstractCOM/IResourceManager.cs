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
        /// <summary>TODO</summary>
        [DispId(2)]
        object LoadImage(string name);
    }
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(RubberduckGuid.IRibbonWorkbook)]
    public interface IRibbonWorkbook {
        /// <summary>TODO</summary>
        [DispId(1)]
        RibbonViewModel ViewModel { get; }
    }
}
