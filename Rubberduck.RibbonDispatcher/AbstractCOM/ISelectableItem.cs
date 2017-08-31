////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////

using System;
using System.Runtime.InteropServices;

namespace Rubberduck.RibbonDispatcher.AbstractCOM {
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(RubberduckGuid.ISelectableItem)]
    public interface ISelectableItem {
        /// <summary>TODO</summary>
        [DispId(DispIds.ItemId)]
        string   Id         { get; }
        /// <summary>TODO</summary>
        [DispId(DispIds.ItemLabel)]
        string   Label      { get; }
        /// <summary>TODO</summary>
        [DispId(DispIds.ItemScreenTip)]
        string   ScreenTip  { get; }
        /// <summary>TODO</summary>
        [DispId(DispIds.ItemSuperTip)]
        string   SuperTip   { get; }

        /// <summary>Call back for GetItemScreenTip events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.ItemShowImage)]
        bool     ShowImage  { get; set; }
        /// <summary>Call back for GetItemSuperTip events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.ItemShowLabel)]
        bool     ShowLabel  { get; set; }
    }
}
