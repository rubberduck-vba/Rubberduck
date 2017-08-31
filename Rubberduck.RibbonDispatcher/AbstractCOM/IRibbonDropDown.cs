////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;

using Rubberduck.RibbonDispatcher.Abstract;

namespace Rubberduck.RibbonDispatcher.AbstractCOM {
    /// <summary>The total interface (required to be) exposed externally by RibbonDropDown objects; 
    /// composition of IRibbonCommon, IDropDownItem &amp; IImageableItem</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(RubberduckGuid.IRibbonDropDown)]
    public interface IRibbonDropDown : IRibbonCommon, IDropDownItem {
        /// <summary>Returns the unique (within this ribbon) identifier for this control.</summary>
        [DispId(DispIds.Id)]
        new string        Id            { get; }
        /// <summary>Only applicable for Menu Items.</summary>
        [DispId(DispIds.Description)]
        new string        Description   { get; }
        /// <summary>TODO</summary>
        [DispId(DispIds.KeyTip)]
        new string        KeyTip        { get; }
        /// <summary>TODO</summary>
        [DispId(DispIds.Label)]
        new string        Label         { get; }
        /// <summary>TODO</summary>
        [DispId(DispIds.ScreenTip)]
        new string        ScreenTip     { get; }
        /// <summary>TODO</summary>
        [DispId(DispIds.SuperTip)]
        new string        SuperTip      { get; }
        /// <summary>Sets the Label, KeyTip, ScreenTip and SuperTip for this control from the supplied values.</summary>
        [DispId(DispIds.SetLanguageStrings)]
        new void          SetLanguageStrings(IRibbonTextLanguageControl languageStrings);

        /// <summary>TODO</summary>
        [DispId(DispIds.IsEnabled)]
        new bool          IsEnabled     { get; set; }
        /// <summary>TODO</summary>
        [DispId(DispIds.IsVisible)]
        new bool          IsVisible     { get; set; }
        /// <summary>TODO</summary>
        [DispId(DispIds.Size)]
        new RdControlSize Size          { get; set; }

        /// <summary>TODO</summary>
        [DispId(DispIds.SelectedItemId)]
        new string      SelectedItemId      { get; set; }
        /// <summary>TODO</summary>
        [DispId(DispIds.SelectedItemIndex)]
        new int         SelectedItemIndex   { get; set; }
        /// <summary>Call back for OnAction events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.OnActionDropDown)]
        new void        OnActionDropDown(string selectedId, int selectedIndex);

        /// <summary>Call back for ItemCount events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.ItemCount)]
        new int         ItemCount           { get; }
        /// <summary>Call back for GetItemID events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.ItemId)]
        new string      ItemId(int index);
        /// <summary>Call back for GetItemLabel events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.ItemLabel)]
        new string      ItemLabel(int index);
        /// <summary>Call back for GetItemScreenTip events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.ItemScreenTip)]
        new string      ItemScreenTip(int index);
        /// <summary>Call back for GetItemSuperTip events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.ItemSuperTip)]
        new string      ItemSuperTip(int index);

        /// <summary>Call back for GetItemSuperTip events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.ItemShowImage)]
        new bool        ItemShowImage(int index);
        /// <summary>Call back for GetItemSuperTip events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.ItemShowLabel)]
        new bool        ItemShowLabel(int index);
    }
}
