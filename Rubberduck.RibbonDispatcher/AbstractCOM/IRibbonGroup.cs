////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;

using Rubberduck.RibbonDispatcher.Abstract;

namespace Rubberduck.RibbonDispatcher.AbstractCOM {
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [Guid("CDC8AF57-3837-4883-906B-7A670BF07711")]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonGroup {
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
        /// <summary>Sets the Label, KeyTip, ScreenTip and SuperTip for this control from the supplied values.</summary>
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
