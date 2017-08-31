////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;
using stdole;

using Rubberduck.RibbonDispatcher.Abstract;

namespace Rubberduck.RibbonDispatcher.AbstractCOM {
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [Guid("D03E9DE1-F37D-40D6-89D6-A6B76A608D97")]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonToggleButton : IRibbonCommon, IToggleItem, IImageableItem {
        /// <summary>TODO</summary>
        new string        Id          { get; }
        /// <summary>Only applicable for Menu Items.</summary>
        new string        Description { get; }
        /// <summary>TODO</summary>
        new string        KeyTip      { get; }
        /// <summary>TODO</summary>
        new string        Label       { get; }
        /// <summary>TODO</summary>
        new string        ScreenTip   { get; }
        /// <summary>TODO</summary>
        new string        SuperTip    { get; }
 
        /// <summary>TODO</summary>
        new object        Image       { get; }
        /// <summary>TODO</summary>
        new bool          IsEnabled   { get; set; }
        /// <summary>TODO</summary>
        new bool          IsPressed   { get; }
        /// <summary>TODO</summary>
        new bool          IsVisible   { get; set; }
        /// <summary>Returns or set whether to show the control's image; ignored by Large controls.</summary>
        new bool          ShowImage   { get; set; }
        /// <summary>Returns or set whether to show the control's label; ignored by Large controls.</summary>
        new bool          ShowLabel   { get; set; }
        /// <summary>TODO</summary>
        new RdControlSize Size        { get; set; }

        /// <summary>TODO</summary>
        new void OnAction(bool isPressed);

        /// <summary>TODO</summary>
        new void SetLanguageStrings(IRibbonTextLanguageControl languageStrings);
        /// <summary>TODO</summary>
        new void SetImage(IPictureDisp image);
        /// <summary>TODO</summary>
        new void SetImageMso(string imageMso);
    }
}
