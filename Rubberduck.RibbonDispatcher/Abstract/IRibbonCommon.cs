using System;
using System.Runtime.InteropServices;

using stdole;
using Microsoft.Office.Core;

namespace Rubberduck.RibbonDispatcher.Abstract
{
    using ControlSize         = RibbonControlSize;
    using ChangedEventHandler = EventHandler<ChangedControlEventArgs>;

    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonCommon {
        string       Id                { get; }

        string       Description       { get; }
        string       KeyTip            { get; }
        string       Label             { get; }
        string       ScreenTip         { get; }
        string       SuperTip          { get; }

        bool         Enabled           { get; set; }
        IPictureDisp Image             { get; set; } 
        ControlSize  Size              { get; set; }
        bool         Visible           { get; set; }

        bool         ShowLabel         { get; set; }
        bool         ShowImage         { get; set; }

        void SetLanguageStrings(IRibbonTextLanguageControl languageStrings);

        event ChangedEventHandler Changed;

        void         OnChanged();
    }
}
