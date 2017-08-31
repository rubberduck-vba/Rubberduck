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
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(RubberduckGuid.IRibbonDropDown)]
    public interface IRibbonDropDown : IRibbonCommon, IDropDownItem, IImageableItem {
        /// <summary>Returns the unique (within this ribbon) identifier for this control.</summary>
        [DispId(DispIds.Id)]
        new string        Id          { get; }
        /// <summary>Only applicable for Menu Items.</summary>
        [DispId(DispIds.Description)]
        new string        Description { get; }
        /// <summary>TODO</summary>
        [DispId(DispIds.KeyTip)]
        new string        KeyTip      { get; }
        /// <summary>TODO</summary>
        [DispId(DispIds.Label)]
        new string        Label       { get; }
        /// <summary>TODO</summary>
        [DispId(DispIds.ScreenTip)]
        new string        ScreenTip   { get; }
        /// <summary>TODO</summary>
        [DispId(DispIds.SuperTip)]
        new string        SuperTip    { get; }
        /// <summary>Sets the Label, KeyTip, ScreenTip and SuperTip for this control from the supplied values.</summary>
        [DispId(DispIds.SetLanguageStrings)]
        new void          SetLanguageStrings(IRibbonTextLanguageControl languageStrings);

        /// <summary>TODO</summary>
        [DispId(DispIds.IsEnabled)]
        new bool          IsEnabled   { get; set; }
        /// <summary>TODO</summary>
        [DispId(DispIds.IsVisible)]
        new bool          IsVisible   { get; set; }
        /// <summary>TODO</summary>
        [DispId(DispIds.Size)]
        new RdControlSize Size        { get; set; }

        /// <summary>TODO</summary>
        [DispId(DispIds.SelectedItemId)]
        new string      SelectedItemId  { get; set; }
        /// <summary>TODO</summary>
        [DispId(DispIds.OnActionDropDown)]
        new void        OnActionDropDown(string itemId);

        /// <summary>TODO</summary>
        [DispId(DispIds.Image)]
        new object        Image       { get; }
        /// <summary>Returns or set whether to show the control's image; ignored by Large controls.</summary>
        [DispId(DispIds.ShowImage)]
        new bool          ShowImage   { get; set; }
        /// <summary>Returns or set whether to show the control's label; ignored by Large controls.</summary>
        [DispId(DispIds.ShowLabel)]
        new bool          ShowLabel   { get; set; }
        /// <summary>TODO</summary>
        [DispId(DispIds.SetImage)]
        new void          SetImage(IPictureDisp image);
        /// <summary>TODO</summary>
        [DispId(DispIds.SetImageMso)]
        new void          SetImageMso(string imageMso);
    }
}
