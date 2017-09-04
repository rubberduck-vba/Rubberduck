////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;

using Rubberduck.RibbonDispatcher.AbstractCOM;

namespace Rubberduck.RibbonDispatcher.EventHandlers {
    /// <summary>TODO</summary>
    [Serializable]
    [CLSCompliant(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IToggledEventArgs))]
    public class ToggledEventArgs : EventArgs, IToggledEventArgs {
        /// <summary>TODO</summary>
        public ToggledEventArgs(bool isPressed) { IsPressed = isPressed; }
        /// <summary>TODO</summary>
        public bool IsPressed   { get; }
    }
}
