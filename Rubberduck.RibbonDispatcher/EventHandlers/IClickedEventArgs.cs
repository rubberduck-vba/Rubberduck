////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;

namespace Rubberduck.RibbonDispatcher.EventHandlers {
    /// <summary>TODO</summary>
    [CLSCompliant(true)]
    public interface IClickedEventArgs {
        /// <summary>TODO</summary>
        [DispId(DispIds.ControlId)]
        int ControlId { get; }
    }
}
