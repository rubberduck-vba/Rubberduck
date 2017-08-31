////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;

namespace Rubberduck.RibbonDispatcher.AbstractCOM {
    /// <summary>The total interface (required to be) exposed externally by RibbonDropDown objects; 
    /// composition of IRibbonCommon, IDropDownItem &amp; IImageableItem</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(RubberduckGuid.IRibbonDropDown)]
    public interface IRibbonDropDown {
        /// <summary>Returns the unique (within this ribbon) identifier for this control.</summary>
        [DispId(DispIds.Id)]
        string        Id            { get; }
        /// <summary>Only applicable for Menu Items.</summary>
        [DispId(DispIds.Description)]
        string        Description   { get; }
        /// <summary>TODO</summary>
        [DispId(DispIds.KeyTip)]
        string        KeyTip        { get; }
        /// <summary>TODO</summary>
        [DispId(DispIds.Label)]
        string        Label         { get; }
        /// <summary>TODO</summary>
        [DispId(DispIds.ScreenTip)]
        string        ScreenTip     { get; }
        /// <summary>TODO</summary>
        [DispId(DispIds.SuperTip)]
        string        SuperTip      { get; }
        /// <summary>Sets the Label, KeyTip, ScreenTip and SuperTip for this control from the supplied values.</summary>
        [DispId(DispIds.SetLanguageStrings)]
        void          SetLanguageStrings(IRibbonTextLanguageControl languageStrings);

        /// <summary>TODO</summary>
        [DispId(DispIds.IsEnabled)]
        bool          IsEnabled     { get; set; }
        /// <summary>TODO</summary>
        [DispId(DispIds.IsVisible)]
        bool          IsVisible     { get; set; }

        /// <summary>TODO</summary>
        [DispId(DispIds.SelectedItemId)]
        string      SelectedItemId      { get; set; }
        /// <summary>TODO</summary>
        [DispId(DispIds.SelectedItemIndex)]
        int         SelectedItemIndex   { get; set; }
        /// <summary>Call back for OnAction events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.OnActionDropDown)]
        void        OnActionDropDown(string selectedId, int selectedIndex);

        /// <summary>Call back for ItemCount events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.ItemCount)]
        int         ItemCount           { get; }
        /// <summary>Call back for GetItemID events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.ItemId)]
        string      ItemId(int index);
        /// <summary>Call back for GetItemLabel events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.ItemLabel)]
        string      ItemLabel(int index);
        /// <summary>Call back for GetItemScreenTip events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.ItemScreenTip)]
        string      ItemScreenTip(int index);
        /// <summary>Call back for GetItemSuperTip events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.ItemSuperTip)]
        string      ItemSuperTip(int index);

        /// <summary>Call back for GetItemSuperTip events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.ItemShowImage)]
        bool        ItemShowImage(int index);
        /// <summary>Call back for GetItemSuperTip events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.ItemShowLabel)]
        bool        ItemShowLabel(int index);
    }
}
