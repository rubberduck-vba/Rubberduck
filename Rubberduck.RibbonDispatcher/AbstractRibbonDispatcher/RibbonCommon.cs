using System;
using stdole;
using Microsoft.Office.Core;

namespace RubberDuck.RibbonDispatcher {
    using ControlSize         = RibbonControlSize;
    using LanguageStrings     = IRibbonTextLanguageControl;
    using ChangedEventHandler = EventHandler<ChangedControlEventArgs>;

    public class RibbonCommon : IRibbonCommon {
        internal RibbonCommon(string id, LanguageStrings strings, bool visible, bool enabled, ControlSize size) {
            Id               = id;
            _languageStrings = strings;
            _visible         = visible;
            _enabled         = enabled;
            _size            = size;
        }

        public event ChangedEventHandler Changed;

        public string       Id          { get; }

        public string       Description => _languageStrings?.Description??"";
        public string       KeyTip      => _languageStrings?.KeyTip??"";
        public string       Label       => Use2ndLabel ? _languageStrings?.AlternateLabel??Id 
                                                       : _languageStrings.Label??Id;
        public string       ScreenTip   => _languageStrings?.ScreenTip??Id;
        public string       SuperTip    => _languageStrings?.SuperTip??"";

        public bool         Enabled     { get {return _enabled;}     set {_enabled     = value; OnChanged();} } private bool         _enabled;
        public IPictureDisp Image       { get {return _image;}       set {_image       = value; OnChanged();} } private IPictureDisp _image;
        public ControlSize  Size        { get {return _size;}        set {_size        = value; OnChanged();} } private ControlSize  _size;
        public bool         Use2ndLabel { get {return _use2ndLabel;} set {_use2ndLabel = value; OnChanged();} } private bool         _use2ndLabel;
        public bool         Visible     { get {return _visible;}     set {_visible     = value; OnChanged();} } private bool         _visible;

        public void SetText(LanguageStrings languageStrings) {
            _languageStrings = languageStrings;
            OnChanged();
        }

        public void OnChanged() => Changed?.Invoke(this, new ChangedControlEventArgs(Id));

        private LanguageStrings _languageStrings;
    }
}
