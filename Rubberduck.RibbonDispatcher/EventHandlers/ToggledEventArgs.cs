using System;
using System.Runtime.InteropServices;

using Rubberduck.RibbonDispatcher.Abstract;

namespace Rubberduck.RibbonDispatcher.EventHandlers {
    /// <summary>TODO</summary>
    [Serializable]
    [CLSCompliant(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class ToggledEventArgs : EventArgs, IToggledEventArgs {
        /// <summary>TODO</summary>
        public ToggledEventArgs(bool isPressed) { IsPressed = isPressed; }
        /// <summary>TODO</summary>
        public bool IsPressed   { get; }
    }
}
