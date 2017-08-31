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
    public class SelectionMadeEventArgs : EventArgs, ISelectionMadeEventArgs {
        /// <summary>TODO</summary>
        public SelectionMadeEventArgs(string itemId) { ItemId = itemId; }
        /// <summary>TODO</summary>
        public string ItemId    { get; }
    }
}
