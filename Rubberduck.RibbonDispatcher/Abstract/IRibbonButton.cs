using System;
using System.Runtime.InteropServices;

namespace Rubberduck.RibbonDispatcher.Abstract
{
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [Guid("EBC076A1-922E-46B7-91D4-A18DF10ABC70")]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonButton : IRibbonCommon, IRibbonImageable {
        /// <summary>TODO</summary>
        void OnAction();
    }
}
