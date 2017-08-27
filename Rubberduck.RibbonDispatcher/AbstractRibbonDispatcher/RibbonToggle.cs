using System;
using Microsoft.Office.Core;

namespace RubberDuck.RibbonDispatcher {

    using ClickedEventHandler = EventHandler<ClickedEventArgs>;
    public class RibbonToggle : RibbonCommon, IRibbonToggle {
        internal RibbonToggle(string id, bool visible, bool enabled, RibbonControlSize size)
            : base(id, visible, enabled, size){
        }

        public event ClickedEventHandler Clicked;

        public bool ShowLabel { get; }
        public bool ShowImage { get; }
        public bool IsPressed { get; set; }

        public void OnAction(bool isPressed) {
            Clicked?.Invoke(this,new ClickedEventArgs(isPressed));
            Use2ndLabel = IsPressed = isPressed;
            OnChanged();
        }

        public IRibbonCommon AsRibbonControl => this;
   }
}
