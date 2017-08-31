////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;

namespace Rubberduck.RibbonDispatcher.AbstractCOM {
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [Guid("2D536C8F-324B-4013-B00C-25608948E416")]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonTextLanguageControl {
        /// <summary>TODO</summary>
        [DispId(1)]
        string Label            { get; }
        /// <summary>TODO</summary>
        [DispId(2)]
        string KeyTip           { get; }
        /// <summary>TODO</summary>
        [DispId(3)]
        string ScreenTip        { get; }
        /// <summary>TODO</summary>
        [DispId(4)]
        string SuperTip         { get; }
        /// <summary>TODO</summary>
        [DispId(5)]
        string AlternateLabel   { get; }
        /// <summary>TODO</summary>
        [DispId(6)]
        string Description      { get; }
    }
}
