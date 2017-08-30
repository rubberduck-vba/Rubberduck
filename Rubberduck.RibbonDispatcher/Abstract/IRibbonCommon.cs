using System;
using System.Runtime.InteropServices;

using stdole;
using Microsoft.Office.Core;

using Rubberduck.RibbonDispatcher.EventHandlers;

namespace Rubberduck.RibbonDispatcher.Abstract
{
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [Guid("1512D081-66D6-49BB-BED1-A25BDDEB5F7F")]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonCommon {
        /// <summary>TODO</summary>
        string            Id                { get; }

        /// <summary>TODO</summary>
        string            Description       { get; }
        /// <summary>TODO</summary>
        string            KeyTip            { get; }
        /// <summary>TODO</summary>
        string            Label             { get; }
        /// <summary>TODO</summary>
        string            ScreenTip         { get; }
        /// <summary>TODO</summary>
        string            SuperTip          { get; }

        /// <summary>TODO</summary>
        bool              IsEnabled         { get; set; }
        /// <summary>TODO</summary>
        IPictureDisp      Image             { get; set; }
        /// <summary>TODO</summary>
        MyRibbonControlSize Size               { get; set; }
        /// <summary>TODO</summary>
        bool              IsVisible          { get; set; }

        /// <summary>TODO</summary>
        bool              ShowLabel          { get; set; }
        /// <summary>TODO</summary>
        bool              ShowImage         { get; set; }

        /// <summary>TODO</summary>
        void SetLanguageStrings(IRibbonTextLanguageControl languageStrings);

        /// <summary>TODO</summary>
        event ChangedEventHandler Changed;

        /// <summary>TODO</summary>
        void OnChanged();
    }
}
