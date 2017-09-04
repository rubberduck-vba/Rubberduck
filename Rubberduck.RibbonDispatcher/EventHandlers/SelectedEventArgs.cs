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
    [ComDefaultInterface(typeof(ISelectedEventArgs))]
    public class SelectedEventArgs : EventArgs, ISelectedEventArgs {
        /// <summary>TODO</summary>
        public SelectedEventArgs(string itemId, int itemIndex) { ItemId = itemId; ItemIndex = itemIndex; }
        /// <summary>TODO</summary>
        public string ItemId    { get; }
        /// <summary>TODO</summary>
        public int    ItemIndex { get; }
    }
}
