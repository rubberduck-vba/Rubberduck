using System;
using System.Runtime.InteropServices;

namespace Rubberduck.RibbonDispatcher.Abstract {
    using ClickedEventHandler = EventHandler<ClickedEventArgs>;
    [ComVisible(true)][CLSCompliant(true)]
    public interface IRibbonToggle {
        event ClickedEventHandler Clicked;

        bool ShowLabel { get; }
        bool ShowImage { get; }
        bool IsPressed { get; }

        void OnAction(bool isPressed);

        IRibbonCommon AsRibbonControl { get; }
    }
}
