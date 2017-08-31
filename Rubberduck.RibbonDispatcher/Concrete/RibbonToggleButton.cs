////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Resources;
using System.Runtime.InteropServices;
using System.Diagnostics.CodeAnalysis;
using stdole;

using Rubberduck.RibbonDispatcher.Abstract;
using Rubberduck.RibbonDispatcher.AbstractCOM;
using Rubberduck.RibbonDispatcher.EventHandlers;

namespace Rubberduck.RibbonDispatcher.Concrete {
    /// <summary>The ViewModel for Ribbon ToggleButton objects.</summary>
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable",
       Justification = "Publc, Non-Creatable class with exported Events.")]
    [Serializable]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IToggledEvents))]
    [ComDefaultInterface(typeof(IRibbonToggleButton))]
    [Guid(RubberduckGuid.RibbonToggleButton)]
    public class RibbonToggleButton : RibbonCommon, IRibbonToggleButton {
        internal RibbonToggleButton(string id, ResourceManager mgr, bool visible, bool enabled, RdControlSize size,
                string imageMso, bool showImage, bool showLabel, ToggledEventHandler onToggledAction)
            : base(id, mgr, visible, enabled, size) {
            _image     = new ImageObject(imageMso);
            _showImage = showImage;
            _showLabel = showLabel;
            if (onToggledAction != null) Toggled += onToggledAction;
        }
        internal RibbonToggleButton(string id, ResourceManager mgr, bool visible, bool enabled, RdControlSize size,
                IPictureDisp image, bool showImage, bool showLabel, ToggledEventHandler onToggledAction)
            : base(id, mgr, visible, enabled, size) {
            _image     = new ImageObject(image);
            _showImage = showImage;
            _showLabel = showLabel;
            if (onToggledAction != null) Toggled += onToggledAction;
        }

        /// <summary>TODO</summary>
        public event ToggledEventHandler Toggled;

        /// <summary>TODO</summary>
        public void OnActionToggle(bool isPressed) {
            IsPressed = isPressed;
            Toggled?.Invoke(this, new ToggledEventArgs(isPressed));
            OnChanged();
        }

        /// <summary>TODO</summary>
        public bool       IsPressed { get; private set; }
        /// <summary>TODO</summary>
        public new string Label       => IsPressed && ! String.IsNullOrEmpty(LanguageStrings?.AlternateLabel ?? "")
                                       ? LanguageStrings?.AlternateLabel??Id 
                                       : LanguageStrings?.Label??Id;

        #region IImageableItem implementation
        /// <inheritdoc/>
        public object Image => _image.Image;
        private ImageObject _image;
        /// <inheritdoc/>
        public bool ShowLabel {
            get { return _showLabel; }
            set { _showLabel = value; OnChanged(); }
        }
        private bool _showLabel;
        /// <inheritdoc/>
        public bool ShowImage {
            get { return _showImage && Image != null; }
            set { _showImage = value; OnChanged(); }
        }
        private bool _showImage;

        /// <inheritdoc/>
        public void SetImage(IPictureDisp image) { _image = new ImageObject(image);     OnChanged(); }
        /// <inheritdoc/>
        public void SetImageMso(string imageMso) { _image = new ImageObject(imageMso); OnChanged(); }
        #endregion
    }

}
