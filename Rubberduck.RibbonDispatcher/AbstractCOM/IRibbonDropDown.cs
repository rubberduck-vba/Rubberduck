////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;
using stdole;

using Rubberduck.RibbonDispatcher.Abstract;
namespace Rubberduck.RibbonDispatcher.AbstractCOM {
    /// <summary>The total interface (required to be) exposed externally by RibbonDropDown objects; 
    /// composition of IRibbonCommon, IDropDownItem &amp; IImageableItem</summary>
    [ComVisible(true)]
    [Guid("7660882A-351B-4518-AFD3-8CA1E3EFE9D8")]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonDropDown : IRibbonCommon, IDropDownItem, IImageableItem {
        /// <summary>Returns the unique (within this ribbon) identifier for this control.</summary>
        [DispId( 1)]
        new string      Id              { get; }
        /// <summary>Only applicable for Menu Items.</summary>
        [DispId( 2)]
        new string      Description     { get; }
        /// <summary>Returns the KeyTip for this control.</summary>
        [DispId( 3)]
        new string      KeyTip          { get; }
        /// <summary>TODO</summary>
        [DispId( 4)]
        new string      Label           { get; }
        /// <summary>TODO</summary>
        [DispId( 5)]
        new string      ScreenTip       { get; }
        /// <summary>TODO</summary>
        [DispId( 6)]
        new string      SuperTip        { get; }
        /// <summary>TODO</summary>
        [DispId( 7)]
        new void        SetLanguageStrings(IRibbonTextLanguageControl languageStrings);

        /// <summary>TODO</summary>
        [DispId( 8)]
        new bool        IsEnabled       { get; set; }
        /// <summary>TODO</summary>
        [DispId( 9)]
        new bool        IsVisible       { get; set; }
        /// <summary>TODO</summary>
        [DispId(10)]
        new RdControlSize Size            { get; set; }

        /// <summary>TODO</summary>
        [DispId(24)]
        new string      SelectedItemId  { get; set; }
        /// <summary>TODO</summary>
        [DispId(25)]
        new void        OnActionDropDown(string itemId);

        /// <summary>TODO</summary>
        [DispId(31)]
        new object      Image       { get; }
        /// <summary>Returns or set whether to show the control's image; ignored by Large controls.</summary>
        [DispId(32)]
        new bool        ShowImage   { get; set; }
        /// <summary>Returns or set whether to show the control's label; ignored by Large controls.</summary>
        [DispId(33)]
        new bool        ShowLabel   { get; set; }
        /// <summary>TODO</summary>
        [DispId(34)]
        new void        SetImage(IPictureDisp image);
        /// <summary>TODO</summary>
        [DispId(35)]
        new void        SetImageMso(string imageMso);
    }
}
