using System;
using System.Collections.Generic;

namespace RubberDuck.Ribbon {
    public interface IRibbonText : IReadOnlyList<IRibbonTextLanguage> {
        string DefaultLangCode { get; }
    }

    /// <summary>All the control-specific {IRibbonTextLanguageControl} for a language.</summary>
    public interface IRibbonTextLanguage : IReadOnlyList<IRibbonTextLanguageControl> {
        string LangCode         { get; }
    }

    /// <summary>All the language-specific {IRibbonTextLanguageControl} for a ribbon control.</summary>
    public interface IRibbonTextControl : IReadOnlyList<IRibbonTextLanguageControl> {
        string ControlId        { get; }
    }

    public class ClickedEventArgs : EventArgs {
        public ClickedEventArgs(bool isPressed) => IsPressed = isPressed;
        public bool IsPressed {get; }
    }

    public class SelectionMadeEventArgs : EventArgs {
        public SelectionMadeEventArgs(string itemID) => ItemID = itemID;
        public string ItemID {get; }
    }
}
