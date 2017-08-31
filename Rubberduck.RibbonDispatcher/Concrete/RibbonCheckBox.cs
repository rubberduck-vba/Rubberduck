﻿////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;
using System.Diagnostics.CodeAnalysis;

using Rubberduck.RibbonDispatcher.ControlMixins;
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
    [ComDefaultInterface(typeof(IRibbonCheckBox))]
    [Guid(RubberduckGuid.RibbonCheckBox)]
    public class RibbonCheckBox : RibbonCommon, IRibbonCheckBox, IToggleableMixin {
        internal RibbonCheckBox(string itemId, IResourceManager mgr, bool visible, bool enabled)
            : base(itemId, mgr, visible, enabled) {
        }

        #region IToggleableDecoration
        /// <summary>TODO</summary>
        public event ToggledEventHandler Toggled;

        /// <summary>TODO</summary>
        public bool IsPressed         { get; private set; }
        /// <summary>TODO</summary>
        public new string Label       => IsPressed && ! string.IsNullOrEmpty(LanguageStrings?.AlternateLabel)
                                       ? LanguageStrings?.AlternateLabel??Id 
                                       : LanguageStrings?.Label??Id;

        /// <summary>TODO</summary>
        public void OnActionToggle(bool isPressed) {
            IsPressed         = isPressed;
            Toggled?.Invoke(this,new ToggledEventArgs(isPressed));
            OnChanged();
        }
        #endregion
   }
}
