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
    public class ClickedEventArgs : EventArgs, IClickedEventArgs {
        /// <summary>TODO</summary>
        public ClickedEventArgs() { ; }
    }
}
