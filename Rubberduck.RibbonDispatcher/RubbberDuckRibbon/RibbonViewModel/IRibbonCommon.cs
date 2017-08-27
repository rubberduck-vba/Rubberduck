using System;

using stdole;
using Microsoft.Office.Core;
using System.Runtime.InteropServices;

namespace RubberDuck.Ribbon {
    [ComVisible(true)][CLSCompliant(true)]
    public interface IRibbonCommon {
        string            ID                { get; }
        string            Description       { get; }
        bool              Enabled           { get; set; }
        IPictureDisp      Image             { get; }
        string            ImageMso          { get; }
        string            KeyTip            { get; }
        string            Label             { get; }
        RibbonControlSize Size              { get; }
        string            ScreenTip         { get; }
        string            SuperTip          { get; }
        bool              UseAlternateLabel { get; set; }
        bool              Visible           { get; set; }

        void              SetText(IRibbonTextLanguageControl strings);

        void InitializeImage(IPictureDisp image);
        void InitializeImageMso(string imageMso);

        event EventHandler Changed;
    }
}
