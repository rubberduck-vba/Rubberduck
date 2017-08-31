using System;
using System.Runtime.InteropServices;
using stdole;

namespace Rubberduck.RibbonDispatcher.Abstract {
    [ComVisible(true)]
    [Guid("42D56042-3FE9-4F1F-AD49-3ED0EE6CC987")]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonImageable {
        /// <summary>TODO</summary>
        IPictureDisp Image     { get; set; }
        /// <summary>TODO</summary>
        bool         ShowImage { get; set; }
        /// <summary>TODO</summary>
        bool         ShowLabel { get; set; }
    }
}
