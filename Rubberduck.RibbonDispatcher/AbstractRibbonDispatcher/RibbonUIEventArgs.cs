using System;

using Microsoft.Office.Core;

namespace Rubberduck.RibbonDispatcher.Abstract {
    public class RibbonUIEventArgs : EventArgs {
        public RibbonUIEventArgs(IRibbonUI ribbonUI) {
            RibbonUI = ribbonUI;
        }
        public IRibbonUI RibbonUI {get;}
    }
}
