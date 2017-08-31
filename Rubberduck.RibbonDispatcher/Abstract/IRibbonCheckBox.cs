using System;
using System.Runtime.InteropServices;

namespace Rubberduck.RibbonDispatcher.Abstract {

    /// <summary>TODO</summary>
    [ComVisible(true)]
    [Guid("3472EE69-B8D7-44FB-8753-735D9E5D26F1")]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonCheckBox : IRibbonCommon, IToggleItem {
    }
}
