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
    [ComDefaultInterface(typeof(ISelectionMadeEventArgs))]
    public class SelectionMadeEventArgs : EventArgs, ISelectionMadeEventArgs {
        /// <summary>TODO</summary>
        public SelectionMadeEventArgs(string itemId, int itemIndex) { ItemId = itemId; ItemIndex = itemIndex; }
        /// <summary>TODO</summary>
        public string ItemId    { get; }
        /// <summary>TODO</summary>
        public int    ItemIndex { get; }
    }
}
