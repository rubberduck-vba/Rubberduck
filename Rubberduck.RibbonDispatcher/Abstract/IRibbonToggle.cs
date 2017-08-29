using System;
using System.Runtime.InteropServices;

namespace Rubberduck.RibbonDispatcher.Abstract
{
    using ClickedEventHandler = EventHandler<ClickedEventArgs>;

    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonToggle : IRibbonCommon {
        event ClickedEventHandler Clicked;

        bool IsPressed         { get; }
        bool UseAlternateLabel { get; }

        void OnAction(bool isPressed);
    }
}
