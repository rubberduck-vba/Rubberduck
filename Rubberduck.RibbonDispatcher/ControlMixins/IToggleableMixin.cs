////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;

namespace Rubberduck.RibbonDispatcher.ControlMixins {
    /// <summary>The total interface (required to be) exposed externally by RibbonButton objects.</summary>
    [CLSCompliant(true)]
    public interface IToggleableMixin {
        /// <summary>TODO</summary>
        [DispId(DispIds.IsPressed)]
        bool IsPressed { get; }

        /// <summary>TODO</summary>
        [DispId(DispIds.OnActionToggle)]
        void OnActionToggle(bool isPressed);
    }
}
