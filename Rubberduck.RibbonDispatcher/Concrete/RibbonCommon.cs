////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Globalization;
using System.Resources;
using System.Runtime.InteropServices;

using Rubberduck.RibbonDispatcher.AbstractCOM;
using Rubberduck.RibbonDispatcher.EventHandlers;
using LanguageStrings = Rubberduck.RibbonDispatcher.AbstractCOM.IRibbonTextLanguageControl;

namespace Rubberduck.RibbonDispatcher.Concrete {

    /// <summary>TODO</summary>
    [Serializable]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IRibbonCommon))]
    [Guid(RubberduckGuid.RibbonCommon)]
    public abstract class RibbonCommon : IRibbonCommon {
        /// <summary>TODO</summary>
        protected RibbonCommon(string itemId, ResourceManager resourceManager)
            : this(itemId, resourceManager, true, true) {;}
        /// <summary>TODO</summary>
        protected RibbonCommon(string itemId, ResourceManager resourceManager, bool visible, bool enabled) {
            Id               = itemId;
            LanguageStrings  = GetLanguageStrings(itemId, resourceManager);
            _visible         = visible;
            _enabled         = enabled;
        }

        /// <summary>TODO</summary>
        internal event ChangedEventHandler Changed;

        /// <inheritdoc/>
        public string Id          { get; }
        /// <inheritdoc/>
        public string Description => LanguageStrings?.Description ?? "";
        /// <inheritdoc/>
        public string KeyTip      => LanguageStrings?.KeyTip ?? "";
        /// <inheritdoc/>
        public string Label       => LanguageStrings?.Label ?? Id;
        /// <inheritdoc/>
        public string ScreenTip   => LanguageStrings?.ScreenTip ?? Id;
        /// <inheritdoc/>
        public string SuperTip    => LanguageStrings?.SuperTip ?? "";

        /// <inheritdoc/>
        protected LanguageStrings LanguageStrings { get; private set; }

        /// <inheritdoc/>
        public bool IsEnabled {
            get { return _enabled; }
            set { _enabled = value; OnChanged(); }
        }
        private bool _enabled;

        /// <inheritdoc/>
        public bool IsVisible {
            get { return _visible; }
            set { _visible = value; OnChanged(); }
        }
        private bool _visible;

        /// <inheritdoc/>
        public void SetLanguageStrings(LanguageStrings languageStrings) {
            LanguageStrings = languageStrings;
            OnChanged();
        }

        /// <inheritdoc/>
        public void OnChanged() => Changed?.Invoke(this, new ControlChangedEventArgs(Id));

        private static LanguageStrings GetLanguageStrings(string controlId, ResourceManager mgr)
            => new RibbonTextLanguageControl(
                    mgr.GetCurrentUIString(Invariant($"{controlId ?? ""}_Label"))          ?? controlId,
                    mgr.GetCurrentUIString(Invariant($"{controlId ?? ""}_ScreenTip"))      ?? controlId + " SuperTip",
                    mgr.GetCurrentUIString(Invariant($"{controlId ?? ""}_SuperTip"))       ?? controlId + " ScreenTip",
                    mgr.GetCurrentUIString(Invariant($"{controlId ?? ""}_KeyTip"))         ?? controlId,
                    mgr.GetCurrentUIString(Invariant($"{controlId ?? ""}_AlternateLabel")) ?? "",
                    mgr.GetCurrentUIString(Invariant($"{controlId ?? ""}_Description"))    ?? controlId + " Description");
        /// <summary>TODO</summary>
        private static string Invariant(string formattable) => String.Format(formattable, CultureInfo.InvariantCulture);
    }
}
