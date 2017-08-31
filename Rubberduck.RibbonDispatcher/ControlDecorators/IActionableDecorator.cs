////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;

namespace Rubberduck.RibbonDispatcher.ControlDecorators {
    /// <summary>The total interface (required to be) exposed externally by RibbonButton objects.</summary>
    [CLSCompliant(true)]
    public interface IActionableDecorator {
        /// <summary>TODO</summary>
        [DispId(DispIds.OnAction)]
        void OnAction();
    }
}
