////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;

using Rubberduck.RibbonDispatcher.Abstract;

namespace Rubberduck.RibbonDispatcher.AbstractCOM {

    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(RubberduckGuid.IRibbonCheckBox)]
    public interface IRibbonCheckBox : IRibbonCommon, IToggleItem {
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
        [DispId(DispIds.IsPressed)]
        new bool        IsPressed       { get; }
        /// <summary>TODO</summary>
        [DispId(DispIds.OnActionToggle)]
        new void        OnActionToggle(bool isPressed);
    }
}
