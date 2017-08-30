using System;
using System.Runtime.InteropServices;

using Rubberduck.RibbonDispatcher.Abstract;

namespace Rubberduck.RibbonDispatcher.EventHandlers {
    /// <summary>TODO</summary>
    [Serializable]
    [CLSCompliant(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class ClickedEventArgs : EventArgs, IClickedEventArgs {
        /// <summary>TODO</summary>
        public ClickedEventArgs() { ; }
    }
}
