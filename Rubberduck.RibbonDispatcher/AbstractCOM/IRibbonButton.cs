using System;
using System.Runtime.InteropServices;
using stdole;

namespace Rubberduck.RibbonDispatcher.AbstractCOM {
    /// <summary>The total interface (required to be) exposed externally by RibbonButton objects; 
    /// composition of IRibbonCommon, IActionItem &amp; IImageableItem</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(RubberduckGuid.IRibbonButton)]
    public interface IRibbonButton {
        /// <summary>Returns the unique (within this ribbon) identifier for this control.</summary>
        [DispId(DispIds.Id)]
        string        Id          { get; }
        /// <summary>Only applicable for Menu Items.</summary>
        [DispId(DispIds.Description)]
        string        Description { get; }
        /// <summary>TODO</summary>
        [DispId(DispIds.KeyTip)]
        string        KeyTip      { get; }
        /// <summary>TODO</summary>
        [DispId(DispIds.Label)]
        string        Label       { get; }
        /// <summary>TODO</summary>
        [DispId(DispIds.ScreenTip)]
        string        ScreenTip   { get; }
        /// <summary>TODO</summary>
        [DispId(DispIds.SuperTip)]
        string        SuperTip    { get; }
        /// <summary>Sets the Label, KeyTip, ScreenTip and SuperTip for this control from the supplied values.</summary>
        [DispId(DispIds.SetLanguageStrings)]
        void          SetLanguageStrings(IRibbonTextLanguageControl languageStrings);

        /// <summary>TODO</summary>
        [DispId(DispIds.IsEnabled)]
        bool          IsEnabled   { get; set; }
        /// <summary>TODO</summary>
        [DispId(DispIds.IsVisible)]
        bool          IsVisible   { get; set; }

        /// <summary>TODO</summary>
        [DispId(DispIds.Size)]
        RdControlSize Size        { get; set; }

        /// <summary>TODO</summary>
        [DispId(DispIds.OnAction)]
        void          OnAction();

        /// <summary>TODO</summary>
        [DispId(DispIds.Image)]
        object        Image       { get; }
        /// <summary>Returns or set whether to show the control's image; ignored by Large controls.</summary>
        [DispId(DispIds.ShowImage)]
        bool          ShowImage   { get; set; }
        /// <summary>Returns or set whether to show the control's label; ignored by Large controls.</summary>
        [DispId(DispIds.ShowLabel)]
        bool          ShowLabel   { get; set; }
        /// <summary>TODO</summary>
        [DispId(DispIds.SetImage)]
        void          SetImage(IPictureDisp Image);
        /// <summary>TODO</summary>
        [DispId(DispIds.SetImageMso)]
        void          SetImageMso(string ImageMso);
    }
}
