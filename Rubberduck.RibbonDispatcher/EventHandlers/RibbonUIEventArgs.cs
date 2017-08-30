using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace Rubberduck.RibbonDispatcher.EventHandlers {
    /// <summary>TODO</summary>
    [Serializable]
    [CLSCompliant(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class RibbonUIEventArgs : EventArgs {
        /// <summary>TODO</summary>
        public RibbonUIEventArgs(IRibbonUI ribbonUI) {
            RibbonUI = ribbonUI;
        }
        /// <summary>TODO</summary>
        public IRibbonUI RibbonUI {get;}
    }
}
