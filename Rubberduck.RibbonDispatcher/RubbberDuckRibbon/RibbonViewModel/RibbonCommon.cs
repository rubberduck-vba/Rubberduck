using System;
using stdole;
using Microsoft.Office.Core;

namespace RubberDuck.Ribbon {
    public class RibbonCommon : IRibbonCommon {
        internal RibbonCommon(string id, bool visible, bool enabled, RibbonControlSize size) {
            _languageStrings = null;
            ID = id; _visible = visible; _enabled = enabled; _size = size;
        }

        public string            ID                { get; }
        public bool              Enabled           { get {return _enabled;}             set {_enabled = value;           OnChanged();} } private bool _enabled;
        public IPictureDisp      Image             { get {return _image;}               set {_image = value;             OnChanged();} } private IPictureDisp _image;
        public string            ImageMso          { get {return _imageMso;}            set {_imageMso = value;          OnChanged();} } private string _imageMso;
        public RibbonControlSize Size              { get {return _size;}                set {_size = value;              OnChanged();} } private RibbonControlSize _size;
        public bool              UseAlternateLabel { get {return _useAlternateLabel;}   set {_useAlternateLabel = value; OnChanged();} } private bool _useAlternateLabel;
        public bool              Visible           { get {return _visible;}             set {_visible = value;           OnChanged();} } private bool _visible;

        public string            Description       => _languageStrings?.Description??"";
        public string            KeyTip            => _languageStrings?.KeyTip??"";
        public string            Label             => UseAlternateLabel ? _languageStrings?.AlternateLabel??ID : _languageStrings.Label??ID;
        public string            ScreenTip         => _languageStrings?.ScreenTip??ID;
        public string            SuperTip          => _languageStrings?.SuperTip??"";

        private IRibbonTextLanguageControl _languageStrings;

        public event EventHandler Changed;

        protected void OnChanged() => Changed?.Invoke(this, EventArgs.Empty);

        public void InitializeImage(IPictureDisp image)     => Image = image;
        public void InitializeImageMso(string imageMso)     => ImageMso = imageMso;

        public void SetText(IRibbonTextLanguageControl languageStrings) {
            _languageStrings = languageStrings;
            OnChanged();
        }
    }
}
