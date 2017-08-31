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
        protected RibbonCommon(string id, ResourceManager mgr, bool visible, bool enabled, MyRibbonControlSize size)
            : this(id, mgr, visible, enabled, size, false, true) {;}
        /// <summary>TODO</summary>
        protected RibbonCommon(string id, ResourceManager mgr, bool visible, bool enabled, MyRibbonControlSize size,
                bool showImage, bool showLabel) {
            Id               = id;
            LanguageStrings  = GetLanguageStrings(id, mgr);
            _visible         = visible;
            _enabled         = enabled;
            _size            = size;
            _showImage       = showImage;
            _showLabel       = showLabel;
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
        public bool ShowLabel {
            get { return _showLabel; }
            set { _showLabel = value; OnChanged(); }
        }
        private bool _showLabel;
        /// <inheritdoc/>
        public bool ShowImage {
            get { return _showImage; }
            set { _showImage = value; OnChanged(); }
        }
        private bool _showImage;

        /// <inheritdoc/>
        public void SetLanguageStrings(LanguageStrings languageStrings) {
            LanguageStrings = languageStrings;
            OnChanged();
        }

        /// <inheritdoc/>
        public void OnChanged() => Changed?.Invoke(this, new ControlChangedEventArgs(Id));

        private static LanguageStrings GetLanguageStrings(string controlId, ResourceManager mgr)
            => new RibbonTextLanguageControl(
                    mgr.GetCurrentUItString($"{controlId ?? ""}_Label")          ?? controlId,
                    mgr.GetCurrentUItString($"{controlId ?? ""}_ScreenTip")      ?? controlId + " SuperTip",
                    mgr.GetCurrentUItString($"{controlId ?? ""}_SuperTip")       ?? controlId + " ScreenTip",
                    mgr.GetCurrentUItString($"{controlId ?? ""}_KeyTip")         ?? controlId,
                    mgr.GetCurrentUItString($"{controlId ?? ""}_AlternateLabel") ?? "",
                    mgr.GetCurrentUItString($"{controlId ?? ""}_Description")    ?? controlId + " Description");
    }

    /// <summary>TODO</summary>
    public static partial class ResourceManagerExtensions {
        /// <summary>TODO</summary>
        public static string GetCurrentUItString(this ResourceManager resourceManager, string name)
            => resourceManager.GetString(name, CultureInfo.CurrentUICulture);
        /// <summary>TODO</summary>
        public static string GetInvariantString(this ResourceManager resourceManager, string name)
            => resourceManager.GetString(name, CultureInfo.InvariantCulture);
    }
}
