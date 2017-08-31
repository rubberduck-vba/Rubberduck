using System;
using System.Resources;
using System.Runtime.InteropServices;

using Rubberduck.RibbonDispatcher.Abstract;
using Rubberduck.RibbonDispatcher.EventHandlers;

namespace Rubberduck.RibbonDispatcher.Concrete {
    /// <summary>TODO</summary>
    [Serializable]
    [CLSCompliant(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IToggledEvents))]
    public class RibbonCheckBox : RibbonCommon, IRibbonToggle {
        internal RibbonCheckBox(string id, ResourceManager mgr, bool visible, bool enabled, MyRibbonControlSize size,
                ToggledEventHandler onToggledAction)
            : base(id, mgr, visible, enabled, size) {
            if (onToggledAction != null) Toggled += onToggledAction;
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
            UseAlternateLabel = isPressed && !String.IsNullOrEmpty(LanguageStrings?.AlternateLabel??"");
            OnChanged();
        }

        /// <summary>TODO</summary>
        public IRibbonCommon AsRibbonControl => this;
   }

}
