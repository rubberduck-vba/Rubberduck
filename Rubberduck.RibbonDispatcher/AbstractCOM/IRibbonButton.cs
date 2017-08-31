////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;
using stdole;

using Rubberduck.RibbonDispatcher.Abstract;

namespace Rubberduck.RibbonDispatcher.AbstractCOM {
    /// <summary>The total interface (required to be) exposed externally by RibbonButton objects; 
    /// composition of IRibbonCommon, IActionItem &amp; IImageableItem</summary>
    [ComVisible(true)]
    [Guid("EBC076A1-922E-46B7-91D4-A18DF10ABC70")]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonButton : IRibbonCommon, IActionItem, IImageableItem {
        /// <summary>Returns the unique (within this ribbon) identifier for this control.</summary>
        [DispId( 1)]
        new string        Id          { get; }
        /// <summary>Only applicable for Menu Items.</summary>
        [DispId( 2)]
        new string        Description { get; }
        /// <summary>TODO</summary>
        [DispId( 3)]
        new string        KeyTip      { get; }
        /// <summary>TODO</summary>
        [DispId( 4)]
        new string        Label       { get; }
        /// <summary>TODO</summary>
        [DispId( 5)]
        new string        ScreenTip   { get; }
        /// <summary>TODO</summary>
        [DispId( 6)]
        new string        SuperTip    { get; }
        /// <summary>Sets the Label, KeyTip, ScreenTip and SuperTip for this control from the supplied values.</summary>
        [DispId( 7)]
        new void          SetLanguageStrings(IRibbonTextLanguageControl languageStrings);

        /// <summary>TODO</summary>
        [DispId( 8)]
        new bool          IsEnabled   { get; set; }
        /// <summary>TODO</summary>
        [DispId( 9)]
        new bool          IsVisible   { get; set; }
        /// <summary>TODO</summary>
        [DispId(10)]
        new RdControlSize Size        { get; set; }

        /// <summary>TODO</summary>
        [DispId(21)]
        new void          OnAction();

        /// <summary>TODO</summary>
        [DispId(31)]
        new object        Image       { get; }
        /// <summary>Returns or set whether to show the control's image; ignored by Large controls.</summary>
        [DispId(32)]
        new bool          ShowImage   { get; set; }
        /// <summary>Returns or set whether to show the control's label; ignored by Large controls.</summary>
        [DispId(33)]
        new bool          ShowLabel   { get; set; }
        /// <summary>TODO</summary>
        [DispId(34)]
        new void          SetImage(IPictureDisp image);
        /// <summary>TODO</summary>
        [DispId(35)]
        new void          SetImageMso(string imageMso);
    }
}
