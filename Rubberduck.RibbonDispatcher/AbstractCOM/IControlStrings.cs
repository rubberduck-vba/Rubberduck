using System;
using System.Runtime.InteropServices;

namespace Rubberduck.RibbonDispatcher.AbstractCOM {
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(RubberduckGuid.IControlStrings)]
    public interface IControlStrings {
        /// <summary>TODO</summary>
        int         Count         { get; }
        /// <summary>TODO</summary>
        string this[string Index] { get; }

        /// <summary>TODO</summary>
        IControlStrings AddControl(string ControlId,
            [Optional]string Label,
            [Optional]string ScreenTip,
            [Optional]string SuperTip,
            [Optional]string AlternateLabel,
            [Optional]string Description,
            [Optional]string KeyTip);
    }
}
