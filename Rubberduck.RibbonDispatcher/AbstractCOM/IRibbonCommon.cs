////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;

namespace Rubberduck.RibbonDispatcher.AbstractCOM {
    /// <summary>The base interface for Ribbnon controls.</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(RubberduckGuid.IRibbonCommon)]
    public interface IRibbonCommon {
        /// <summary>Returns the unique (within this ribbon) identifier for this control.</summary>
        [DispId(DispIds.Id)]
        string Id          { get; }
        /// <summary>Only applicable for Menu Items.</summary>
        [DispId(DispIds.Description)]
        string Description { get; }
        /// <summary>Returns the KeyTip for this control.</summary>
        [DispId(DispIds.KeyTip)]
        string KeyTip      { get; }
        /// <summary>TODO</summary>
        [DispId(DispIds.Label)]
        string Label       { get; }
        /// <summary>TODO</summary>
        [DispId(DispIds.ScreenTip)]
        string ScreenTip   { get; }
        /// <summary>TODO</summary>
        [DispId(DispIds.SuperTip)]
        string SuperTip    { get; }
        /// <summary>TODO</summary>
        [DispId(DispIds.SetLanguageStrings)]
        void   SetLanguageStrings(IRibbonTextLanguageControl languageStrings);

        /// <summary>TODO</summary>
        [DispId(DispIds.IsEnabled)]
        bool   IsEnabled   { get; set; }
        /// <summary>TODO</summary>
        [DispId(DispIds.IsVisible)]
        bool   IsVisible   { get; set; }
    }
}
