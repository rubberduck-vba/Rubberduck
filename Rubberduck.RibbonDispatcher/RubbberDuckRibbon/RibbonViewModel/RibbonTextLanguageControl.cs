using System;
using Microsoft.Office.Core;

namespace RubberDuck.Ribbon {
    public class RibbonTextLanguageControl : IRibbonTextLanguageControl {
        public RibbonTextLanguageControl(
            string label,
            string screentip,
            string supertip,
            string keytip,
            string alternateLabel,
            string description
        ) {
            Label           = label     ?? throw new ArgumentNullException();; 
            ScreenTip       = screentip ?? Label; 
            SuperTip        = supertip  ?? "SuperTip text for " + Label; 
            KeyTip          = keytip    ?? "";
            AlternateLabel  = alternateLabel?? Label; 
            Description     = description   ?? "Description for " + Label;
        }
        public string Label             { get; }
        public string ScreenTip         { get; }
        public string SuperTip          { get; }
        public string KeyTip            { get; }
        public string AlternateLabel    { get; }
        public string Description       { get; }
    }
}
