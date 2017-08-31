////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;

namespace Rubberduck.RibbonDispatcher.EventHandlers {
    /// <summary>Event parameters for a Clicked event.</summary>
    [CLSCompliant(true)]
    public class ClickedEventArgs : EventArgs {
        /// <summary>Returns a new {ClickedEventArgs} instance.</summary>
        public ClickedEventArgs(string controlId) => ControlId = controlId;

        /// <inheritdoc/>
        public string ControlId { get; }
    }
}
