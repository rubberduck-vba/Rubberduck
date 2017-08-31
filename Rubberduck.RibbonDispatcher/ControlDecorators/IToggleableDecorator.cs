////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;

namespace Rubberduck.RibbonDispatcher.ControlDecorators {
    /// <summary>The total interface (required to be) exposed externally by RibbonButton objects.</summary>
    [CLSCompliant(true)]
    public interface IToggleableDecorator {
        /// <summary>TODO</summary>
        [DispId(DispIds.IsPressed)]
        bool IsPressed { get; }

        /// <summary>TODO</summary>
        [DispId(DispIds.OnActionToggle)]
        void OnActionToggle(bool isPressed);
    }
}
