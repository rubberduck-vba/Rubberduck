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
    [Guid(RubberduckGuid.IRibbonCheckBox)]
    public interface IRibbonCheckBox {
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
        [DispId(DispIds.IsPressed)]
        bool        IsPressed       { get; }
        /// <summary>TODO</summary>
        [DispId(DispIds.OnActionToggle)]
        void        OnActionToggle(bool isPressed);
    }
}
