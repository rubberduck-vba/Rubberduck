using System;
using System.Runtime.InteropServices;

namespace Rubberduck.RibbonDispatcher.Abstract {
    /// <summary>The total interface (required to be) exposed externally by RibbonButton objects.</summary>
    [ComVisible(true)]
    [Guid("EBC076A1-922E-46B7-91D4-A18DF10ABC70")]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonButton : IRibbonCommon, IActionItem, IImageableItem { }
}
