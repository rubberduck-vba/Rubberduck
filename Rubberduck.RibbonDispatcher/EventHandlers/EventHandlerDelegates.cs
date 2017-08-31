////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;

namespace Rubberduck.RibbonDispatcher.EventHandlers {
    /// <summary>TODO</summary>
    [SuppressMessage("Microsoft.Design", "CA1003:UseGenericEventHandlerInstances", Justification = "Necessary for COM Interop.")]
    [CLSCompliant(true)]
    public delegate void ToggledEventHandler(object sender, ToggledEventArgs e);

    /// <summary>TODO</summary>
    [SuppressMessage("Microsoft.Design", "CA1003:UseGenericEventHandlerInstances", Justification = "Necessary for COM Interop.")]
    [CLSCompliant(true)]
    public delegate void SelectionMadeEventHandler(object sender, SelectionMadeEventArgs e);

    /// <summary>TODO</summary>
    [SuppressMessage("Microsoft.Design", "CA1003:UseGenericEventHandlerInstances", Justification = "Necessary for COM Interop.")]
    [CLSCompliant(true)]
    public delegate void ClickedEventHandler(object sender, IClickedEventArgs e);
}
