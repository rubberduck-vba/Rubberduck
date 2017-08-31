////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;

using Rubberduck.RibbonDispatcher.Abstract;

namespace Rubberduck.RibbonDispatcher.AbstractCOM {

    /// <summary>TODO</summary>
    [ComVisible(true)]
    [Guid("3472EE69-B8D7-44FB-8753-735D9E5D26F1")]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonCheckBox : IRibbonCommon, IToggleItem {
        /// <summary>Returns the unique (within this ribbon) identifier for this control.</summary>
        [DispId( 1)]
        new string      Id          { get; }
        /// <summary>Only applicable for Menu Items.</summary>
        [DispId( 2)]
        new string      Description { get; }
        /// <summary>TODO</summary>
        [DispId( 3)]
        new string      KeyTip      { get; }
        /// <summary>TODO</summary>
        [DispId( 4)]
        new string      Label       { get; }
        /// <summary>TODO</summary>
        [DispId( 5)]
        new string      ScreenTip   { get; }
        /// <summary>TODO</summary>
        [DispId( 6)]
        new string      SuperTip    { get; }
        /// <summary>TODO</summary>
        [DispId( 7)]
        new void        SetLanguageStrings(IRibbonTextLanguageControl languageStrings);

        /// <summary>TODO</summary>
        [DispId( 8)]
        new bool        IsEnabled   { get; set; }
        /// <summary>TODO</summary>
        [DispId( 9)]
        new bool        IsVisible   { get; set; }
        /// <summary>TODO</summary>
        [DispId(10)]
        new RdControlSize Size        { get; set; }

        /// <summary>TODO</summary>
        [DispId(22)]
        new bool        IsPressed   { get; }
        /// <summary>TODO</summary>
        [DispId(23)]
        new void        OnAction(bool isPressed);
    }
}
