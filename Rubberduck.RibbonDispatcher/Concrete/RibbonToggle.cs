using System;
using System.Runtime.InteropServices;

using Rubberduck.RibbonDispatcher.Abstract;
using Rubberduck.RibbonDispatcher.EventHandlers;

namespace Rubberduck.RibbonDispatcher.Concrete
{
    using LanguageStrings     = IRibbonTextLanguageControl;

    /// <summary>TODO</summary>
    [Serializable]
    [CLSCompliant(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IToggledEvents))]
    public class RibbonToggle : RibbonCommon, IRibbonToggle {
        internal RibbonToggle(string id, LanguageStrings strings, bool visible, bool enabled, MyRibbonControlSize size,
                bool showImage, bool showLabel, ToggledEventHandler onClickedAction)
            : base(id, strings, visible, enabled, size, showImage, showLabel) {
            if (onClickedAction != null) Toggled += onClickedAction;
        }

        /// <summary>TODO</summary>
        public event ToggledEventHandler Toggled;

        /// <summary>TODO</summary>
        public new string Label       => UseAlternateLabel ? LanguageStrings?.AlternateLabel??Id 
                                                           : LanguageStrings?.Label??Id;
        /// <summary>TODO</summary>
        public bool IsPressed         { get; private set; }

        /// <summary>TODO</summary>
        public bool UseAlternateLabel { get; private set; }

        /// <summary>TODO</summary>
        public void OnAction(bool isPressed) {
            Toggled?.Invoke(this,new ToggledEventArgs(isPressed));
            IsPressed         = isPressed;
            UseAlternateLabel = isPressed;
            OnChanged();
        }

        /// <summary>TODO</summary>
        public IRibbonCommon AsRibbonControl => this;
   }

}
