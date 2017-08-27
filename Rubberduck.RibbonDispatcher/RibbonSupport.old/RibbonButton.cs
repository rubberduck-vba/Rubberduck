using System;
using Microsoft.Office.Core;

namespace RubberDuck.RibbonSupport {
    public class RibbonButton : RibbonCommon, IRibbonButton {
        internal RibbonButton(string id, bool visible, bool enabled, RibbonControlSize size)
            : base(id, visible, enabled, size){
        }

        public event EventHandler Clicked;

        public bool ShowLabel { get; set; }
        public bool ShowImage { get; set; }

        public void OnAction() => Clicked?.Invoke(this,null);
 
        public IRibbonCommon AsRibbonControl => this;
  }
}
