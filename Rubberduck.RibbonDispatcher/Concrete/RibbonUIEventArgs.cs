using System;

using Microsoft.Office.Core;

using Rubberduck.RibbonDispatcher;
using Rubberduck.RibbonDispatcher.Abstract;

namespace Rubberduck.RibbonDispatcher.Concrete
{
    public class RibbonUIEventArgs : EventArgs {
        public RibbonUIEventArgs(IRibbonUI ribbonUI) {
            RibbonUI = ribbonUI;
        }
        public IRibbonUI RibbonUI {get;}
    }
}
