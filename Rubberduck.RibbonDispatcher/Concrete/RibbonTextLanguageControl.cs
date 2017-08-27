using System;
using Microsoft.Office.Core;

using Rubberduck.RibbonDispatcher;
using Rubberduck.RibbonDispatcher.Abstract;

namespace Rubberduck.RibbonDispatcher.Concrete
{
    public class RibbonTextLanguageControl : IRibbonTextLanguageControl {
        public RibbonTextLanguageControl(
            string label,
            string screenTip,
            string superTip,
            string keyTip,
            string alternateLabel,
            string description
        ) {
            if (label == null) throw new ArgumentNullException(nameof(label)); 
            Label           = label; 
            ScreenTip       = screenTip     ?? Label; 
            SuperTip        = superTip      ?? "SuperTip text for " + Label; 
            KeyTip          = keyTip        ?? "";
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
