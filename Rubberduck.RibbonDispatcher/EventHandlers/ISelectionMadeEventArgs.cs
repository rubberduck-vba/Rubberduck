﻿////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;
using System.Diagnostics.CodeAnalysis;

namespace Rubberduck.RibbonDispatcher.AbstractCOM {
    /// <summary>TODO</summary>
    [SuppressMessage("Microsoft.Naming", "CA1711:IdentifiersShouldNotHaveIncorrectSuffix", Justification = "Necessary for COM Interop.")]
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(RubberduckGuid.ISelectionMadeEventArgs)]
    public interface ISelectionMadeEventArgs {
        /// <summary>TODO</summary>
        [DispId(DispIds.SelectedItemId)]
        string ItemId    { get; }
        /// <summary>TODO</summary>
        [DispId(DispIds.SelectedItemIndex)]
        int    ItemIndex { get; }
    }
}
