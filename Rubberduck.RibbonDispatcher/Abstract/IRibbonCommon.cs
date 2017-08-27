using System;

using stdole;
using Microsoft.Office.Core;
using System.Runtime.InteropServices;

using Rubberduck.RibbonDispatcher;

namespace Rubberduck.RibbonDispatcher.Abstract
{
    using ControlSize         = RibbonControlSize;
    using ChangedEventHandler = EventHandler<ChangedControlEventArgs>;

    [ComVisible(true)][CLSCompliant(true)]
    public interface IRibbonCommon {
        string       Id            { get; }

        string       Description   { get; }
        string       KeyTip        { get; }
        string       Label         { get; }
        string       ScreenTip     { get; }
        string       SuperTip      { get; }

        bool         Enabled       { get; set; }
        IPictureDisp Image         { get; set; } 
        ControlSize  Size          { get; set; }
        bool         Use2ndLabel   { get; set; }
        bool         Visible       { get; set; }

        void         SetText(IRibbonTextLanguageControl languageStrings);

        event ChangedEventHandler Changed;

        void         OnChanged();
    }
}
