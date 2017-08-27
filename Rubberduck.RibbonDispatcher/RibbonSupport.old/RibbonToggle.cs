using System;
using Microsoft.Office.Core;

namespace RubberDuck.RibbonSupport {

    using ClickedEventHandler = EventHandler<ClickedEventArgs>;
    public class RibbonToggle : RibbonCommon, IRibbonToggle {
        internal RibbonToggle(string id, bool visible, bool enabled, RibbonControlSize size)
            : base(id, visible, enabled, size){
        }

        public event ClickedEventHandler Clicked;

        public bool ShowLabel { get; set; }
        public bool ShowImage { get; set; }
        public bool IsPressed { get; }

        public void OnAction(bool isPressed) {
            Clicked?.Invoke(this,new ClickedEventArgs(isPressed));
            UseAlternateLabel = isPressed;
        }

        public IRibbonCommon AsRibbonControl => this;
   }
}
