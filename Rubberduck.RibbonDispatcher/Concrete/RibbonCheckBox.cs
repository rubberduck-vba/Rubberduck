////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Resources;
using System.Runtime.InteropServices;
using System.Diagnostics.CodeAnalysis;

using Rubberduck.RibbonDispatcher.Abstract;
using Rubberduck.RibbonDispatcher.AbstractCOM;
using Rubberduck.RibbonDispatcher.EventHandlers;

namespace Rubberduck.RibbonDispatcher.Concrete {
    /// <summary>TODO</summary>
    [Serializable]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IToggledEvents))]
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable",
        Justification = "Publc, Non-Creatable class with exported Events.")]
    public class RibbonCheckBox : RibbonCommon, IRibbonCheckBox {
        internal RibbonCheckBox(string id, ResourceManager mgr, bool visible, bool enabled, RdControlSize size,
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
   }

}
