using System;
using System.Runtime.InteropServices;

using Rubberduck.RibbonDispatcher;

namespace Rubberduck.RibbonDispatcher.Abstract
{
    [ComVisible(true)][CLSCompliant(true)]
    public interface IRibbonTextLanguageControl {
        string Label            { get; }
        string AlternateLabel   { get; }
        string KeyTip           { get; }
        string ScreenTip        { get; }
        string SuperTip         { get; }
        string Description      { get; }
    }
}
