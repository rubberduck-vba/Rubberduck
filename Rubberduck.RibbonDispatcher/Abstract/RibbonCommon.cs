using System;
using System.Runtime.InteropServices;

using stdole;
using Microsoft.Office.Core;

namespace Rubberduck.RibbonDispatcher.Abstract
{
    using LanguageStrings     = IRibbonTextLanguageControl;
    using ChangedEventHandler = EventHandler<ChangedControlEventArgs>;

    [ComVisible(true)]
    [CLSCompliant(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public abstract class RibbonCommon : IRibbonCommon {
        protected RibbonCommon(string id, LanguageStrings strings, bool visible, bool enabled, RibbonControlSize size)
            : this(id, strings, visible, enabled, size, false, true) {;}
        protected RibbonCommon(string id, LanguageStrings strings, bool visible, bool enabled, RibbonControlSize size,
                bool showImage, bool showLabel) {
            Id               = id;
            LanguageStrings = strings;
            _visible         = visible;
            _enabled         = enabled;
            _size            = size;
            _showImage       = showImage;
            _showLabel       = showLabel;
        }

        public event ChangedEventHandler Changed;

        public string Id          { get; }
        public string Description => LanguageStrings?.Description ?? "";
        public string KeyTip      => LanguageStrings?.KeyTip ?? "";
        public string Label       => LanguageStrings?.Label ?? Id;
        public string ScreenTip   => LanguageStrings?.ScreenTip ?? Id;
        public string SuperTip    => LanguageStrings?.SuperTip ?? "";

        protected LanguageStrings LanguageStrings { get; private set; }

        public bool Enabled {
            get { return _enabled; }
            set { _enabled = value; OnChanged(); }
        }
        private bool _enabled;

        public IPictureDisp Image {
            get { return _image; }
            set { _image = value; OnChanged(); }
        }
        private IPictureDisp _image;

        public RibbonControlSize Size {
            get { return _size; }
            set { _size = value; OnChanged(); }
        }
        private RibbonControlSize _size;

        public bool Visible {
            get { return _visible; }
            set { _visible = value; OnChanged(); }
        }
        private bool _visible;

        public bool ShowLabel {
            get { return _showLabel; }
            set { _showLabel = value; OnChanged(); }
        }
        private bool _showLabel;
        public bool ShowImage {
            get { return _showImage; }
            set { _showImage = value; OnChanged(); }
        }
        private bool _showImage;

        public void SetLanguageStrings(LanguageStrings languageStrings) {
            LanguageStrings = languageStrings;
            OnChanged();
        }

        public void OnChanged() => Changed?.Invoke(this, new ChangedControlEventArgs(Id));
    }
}
