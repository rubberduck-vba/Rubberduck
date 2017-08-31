////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;

namespace Rubberduck.RibbonDispatcher.EventHandlers {
    /// <summary>Event parameters for a Clicked event.</summary>
    [Serializable]
    [CLSCompliant(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IClickedEventArgs))]
    public class ClickedEventArgs : EventArgs, IClickedEventArgs {
        /// <summary>Returns a new {ClickedEventArgs} instance.</summary>
        public ClickedEventArgs(int controlId) { ControlId = controlId; }

        /// <inheritdoc/>
        public int ControlId { get; }
    }
}
