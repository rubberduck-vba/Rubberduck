////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;

namespace Rubberduck.RibbonDispatcher.EventHandlers {
    /// <summary>TODO</summary>
    [SuppressMessage("Microsoft.Design", "CA1003:UseGenericEventHandlerInstances", Justification = "Necessary for COM Interop.")]
    [CLSCompliant(true)]
    public delegate void ToggledEventHandler(bool IsPressed);

    /// <summary>TODO</summary>
    [SuppressMessage("Microsoft.Design", "CA1003:UseGenericEventHandlerInstances", Justification = "Necessary for COM Interop.")]
    [CLSCompliant(true)]
    public delegate void SelectedEventHandler(object sender, SelectedEventArgs e);

    /// <summary>TODO</summary>
    public delegate void ClickedEventHandler();
}
