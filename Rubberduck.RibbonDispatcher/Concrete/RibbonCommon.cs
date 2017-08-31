using System;
using System.Resources;
using System.Runtime.InteropServices;

using stdole;

using Rubberduck.RibbonDispatcher.Abstract;
using Rubberduck.RibbonDispatcher.EventHandlers;
using System.Globalization;

namespace Rubberduck.RibbonDispatcher.Concrete {
    using LanguageStrings     = IRibbonTextLanguageControl;

    /// <summary>TODO</summary>
    [Serializable]
    [CLSCompliant(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public abstract class RibbonCommon : IRibbonCommon {
        /// <summary>TODO</summary>
        protected RibbonCommon(string id, ResourceManager resourceManager, bool visible, bool enabled, MyRibbonControlSize size) {
            Id               = id;
            LanguageStrings  = GetLanguageStrings(id, resourceManager);
            _visible         = visible;
            _enabled         = enabled;
            _size            = size;
        }

        /// <summary>TODO</summary>
        public event ChangedEventHandler Changed;

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
        public IPictureDisp Image {
            get { return _image; }
            set { _image = value; OnChanged(); }
        }
        private IPictureDisp _image;

        /// <inheritdoc/>
        public MyRibbonControlSize Size {
            get { return _size; }
            set { _size = value; OnChanged(); }
        }
        private MyRibbonControlSize _size;

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
                    mgr.GetCurrentUItString(Invariant($"{controlId ?? ""}_Label"))          ?? controlId,
                    mgr.GetCurrentUItString(Invariant($"{controlId ?? ""}_ScreenTip"))      ?? controlId + " SuperTip",
                    mgr.GetCurrentUItString(Invariant($"{controlId ?? ""}_SuperTip"))       ?? controlId + " ScreenTip",
                    mgr.GetCurrentUItString(Invariant($"{controlId ?? ""}_KeyTip"))         ?? controlId,
                    mgr.GetCurrentUItString(Invariant($"{controlId ?? ""}_AlternateLabel")) ?? "",
                    mgr.GetCurrentUItString(Invariant($"{controlId ?? ""}_Description"))    ?? controlId + " Description");
        /// <summary>TODO</summary>
        public static string Invariant(string formattable) => String.Format(formattable, CultureInfo.InvariantCulture);
    }

    /// <summary>TODO</summary>
    public static partial class ResourceManagerExtensions {
        /// <summary>TODO</summary>
        public static string GetCurrentUItString(this ResourceManager resourceManager, string name)
            => resourceManager?.GetString(name, CultureInfo.CurrentUICulture) ?? "";
        /// <summary>TODO</summary>
        public static string GetInvariantString(this ResourceManager resourceManager, string name)
            => resourceManager?.GetString(name, CultureInfo.InvariantCulture) ?? "";
    }
}
