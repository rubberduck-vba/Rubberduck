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
    [Guid(RubberduckGuid.IRibbonTextLanguageControl)]
    public interface IRibbonTextLanguageControl {
        /// <summary>TODO</summary>
        [DispId(2)]
        string Description      { get; }
        /// <summary>TODO</summary>
        [DispId(3)]
        string Label            { get; }
        /// <summary>TODO</summary>
        [DispId(4)]
        string KeyTip           { get; }
        /// <summary>TODO</summary>
        [DispId(5)]
        string ScreenTip        { get; }
        /// <summary>TODO</summary>
        [DispId(6)]
        string SuperTip         { get; }
        /// <summary>TODO</summary>
        [DispId(7)]
        string AlternateLabel   { get; }
    }
}
