using System;

using Microsoft.Office.Core;

using Rubberduck.RibbonDispatcher.Abstract;

namespace Rubberduck.RibbonDispatcher.Concrete
{
    using LanguageStrings     = IRibbonTextLanguageControl;
    using ClickedEventHandler = EventHandler<ClickedEventArgs>;

    public class RibbonToggle : RibbonCommon, IRibbonToggle {
        internal RibbonToggle(string id, LanguageStrings strings, bool visible, bool enabled, RibbonControlSize size,
                bool showImage, bool showLabel, ClickedEventHandler onClickedAction)
            : base(id, strings, visible, enabled, size, showImage, showLabel) {
            if (onClickedAction != null) Clicked += onClickedAction;
        }

        public event ClickedEventHandler Clicked;

        public new string Label       => UseAlternateLabel ? LanguageStrings?.AlternateLabel??Id 
                                                           : LanguageStrings?.Label??Id;
        public bool IsPressed         { get; private set; }

        public bool UseAlternateLabel { get; private set; }

        public void OnAction(bool isPressed) {
            Clicked?.Invoke(this,new ClickedEventArgs(isPressed));
            IsPressed         = isPressed;
            UseAlternateLabel = isPressed;
            OnChanged();
        }

        public IRibbonCommon AsRibbonControl => this;
   }
}
