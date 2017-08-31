////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;

using Rubberduck.RibbonDispatcher.Abstract;

namespace Rubberduck.RibbonDispatcher.AbstractCOM {
    /// <summary>The base interface for Ribbnon controls.</summary>
    [ComVisible(true)]
    [Guid("1512D081-66D6-49BB-BED1-A25BDDEB5F7F")]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonCommon {
        /// <summary>Returns the unique (within this ribbon) identifier for this control.</summary>
        [DispId( 1)]
        string      Id          { get; }
        /// <summary>Only applicable for Menu Items.</summary>
        [DispId( 2)]
        string      Description { get; }
        /// <summary>Returns the KeyTip for this control.</summary>
        [DispId( 3)]
        string      KeyTip      { get; }
        /// <summary>TODO</summary>
        [DispId( 4)]
        string      Label       { get; }
        /// <summary>TODO</summary>
        [DispId( 5)]
        string      ScreenTip   { get; }
        /// <summary>TODO</summary>
        [DispId( 6)]
        string      SuperTip    { get; }
        /// <summary>TODO</summary>
        [DispId( 7)]
        void        SetLanguageStrings(IRibbonTextLanguageControl languageStrings);

        /// <summary>TODO</summary>
        [DispId( 8)]
        bool        IsEnabled   { get; set; }
        /// <summary>TODO</summary>
        [DispId( 9)]
        bool        IsVisible   { get; set; }
        /// <summary>TODO</summary>
        [DispId(10)]
        RdControlSize Size        { get; set; }
    }
}
