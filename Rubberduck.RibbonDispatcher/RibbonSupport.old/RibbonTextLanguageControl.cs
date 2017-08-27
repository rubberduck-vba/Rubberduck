using Microsoft.Office.Core;

namespace RubberDuck.RibbonSupport {
    public class RibbonTextLanguageControl : IRibbonTextLanguageControl {
        public RibbonTextLanguageControl(
            string label,
            string screentip,
            string supertip,
            string keytip,
            string alternateLabel,
            string description
        ) {
            Label = label; ScreenTip = screentip; SuperTip = supertip; KeyTip = keytip;
            AlternateLabel = alternateLabel; Description = description;
        }
        public string Label             { get; }
        public string ScreenTip         { get; }
        public string SuperTip          { get; }
        public string KeyTip            { get; }
        public string AlternateLabel    { get; }
        public string Description       { get; }
    }
}
