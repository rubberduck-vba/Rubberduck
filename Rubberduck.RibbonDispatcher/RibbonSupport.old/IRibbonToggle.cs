using System;
using System.Runtime.InteropServices;

namespace RubberDuck.RibbonSupport {
    using ClickedEventHandler = EventHandler<ClickedEventArgs>;
    [ComVisible(true)][CLSCompliant(true)]
    public interface IRibbonToggle {
        event ClickedEventHandler Clicked;

        bool ShowLabel { get; set; }
        bool ShowImage { get; set; }
        bool IsPressed { get; }

        void OnAction(bool isPressed);

        IRibbonCommon AsRibbonControl { get; }
    }
}
