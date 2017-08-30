using System;
using System.Runtime.InteropServices;

using Rubberduck.RibbonDispatcher.Abstract;

namespace Rubberduck.RibbonDispatcher.Concrete
{
    /// <summary>TODO</summary>
    [Serializable]
    [CLSCompliant(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class RibbonTextLanguageControl : IRibbonTextLanguageControl {
        /// <summary>TODO</summary>
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
        /// <inheritdoc/>
        public string Label             { get; }
        /// <inheritdoc/>
        public string ScreenTip         { get; }
        /// <inheritdoc/>
        public string SuperTip          { get; }
        /// <inheritdoc/>
        public string KeyTip            { get; }
        /// <inheritdoc/>
        public string AlternateLabel    { get; }
        /// <inheritdoc/>
        public string Description       { get; }
    }
}
