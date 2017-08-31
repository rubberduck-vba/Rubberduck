using System;
using System.Runtime.InteropServices;

namespace Rubberduck.RibbonDispatcher.Abstract {
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [Guid("1512D081-66D6-49BB-BED1-A25BDDEB5F7F")]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonCommon {
        /// <summary>TODO</summary>
        string            Id                { get; }

        /// <summary>Only applicable for Menu Items.</summary>
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
        MyRibbonControlSize Size            { get; set; }
        /// <summary>TODO</summary>
        bool              IsVisible         { get; set; }

        /// <summary>TODO</summary>
        void SetLanguageStrings(IRibbonTextLanguageControl languageStrings);

        /// <summary>TODO</summary>
        void OnChanged();
    }
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [Guid("CDC8AF57-3837-4883-906B-7A670BF07711")]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonGroup {
        /// <summary>TODO</summary>
        string Id { get; }

        /// <summary>TODO</summary>
        string Description { get; }
        /// <summary>TODO</summary>
        string KeyTip { get; }
        /// <summary>TODO</summary>
        string Label { get; }
        /// <summary>TODO</summary>
        string ScreenTip { get; }
        /// <summary>TODO</summary>
        string SuperTip { get; }

        /// <summary>TODO</summary>
        bool IsEnabled { get; set; }
        /// <summary>TODO</summary>
        bool IsVisible { get; set; }

        /// <summary>TODO</summary>
        void SetLanguageStrings(IRibbonTextLanguageControl languageStrings);

        /// <summary>TODO</summary>
        void OnChanged();
    }
}
