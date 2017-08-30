using System;
using System.Runtime.InteropServices;

namespace Rubberduck.RibbonDispatcher.Abstract
{
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [Guid("2D536C8F-324B-4013-B00C-25608948E416")]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonTextLanguageControl {
        /// <summary>TODO</summary>
        string Label            { get; }
        /// <summary>TODO</summary>
        string AlternateLabel   { get; }
        /// <summary>TODO</summary>
        string KeyTip           { get; }
        /// <summary>TODO</summary>
        string ScreenTip        { get; }
        /// <summary>TODO</summary>
        string SuperTip         { get; }
        /// <summary>TODO</summary>
        string Description      { get; }
    }
}
