using System;

using Microsoft.Office.Core;

namespace RubberDuck.RibbonSupport {
    public class RibbonUIEventArgs : EventArgs {
        public RibbonUIEventArgs(IRibbonUI ribbonUI) {
            RibbonUI = ribbonUI;
        }
        public IRibbonUI RibbonUI {get;}
    }
}
