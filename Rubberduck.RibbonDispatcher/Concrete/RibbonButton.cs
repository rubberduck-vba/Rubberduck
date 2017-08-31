////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Resources;
using System.Runtime.InteropServices;
using System.Diagnostics.CodeAnalysis;
using stdole;

using Rubberduck.RibbonDispatcher.ControlDecorators;
using Rubberduck.RibbonDispatcher.AbstractCOM;

namespace Rubberduck.RibbonDispatcher.Concrete {
    /// <summary>The ViewModel for Ribbon Button objects.</summary>
    [Serializable]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IClickedEvents))]
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable",
        Justification = "Publc, Non-Creatable class with exported Events.")]
    [ComDefaultInterface(typeof(IRibbonButton))]
    [Guid(RubberduckGuid.RibbonButton)]
    public class RibbonButton : RibbonCommon, IRibbonButton,
        ISizeableDecorator, IActionableDecorator, IImageableDecorator {
        internal RibbonButton(string itemId, ResourceManager mgr, bool visible, bool enabled, RdControlSize size,
                string imageMso, bool showImage, bool showLabel, EventHandler onClickedAction)
            : base(itemId, mgr, visible, enabled) {
            _size      = size;
            _image     = new ImageObject(imageMso);
            _showImage = showImage;
            _showLabel = showLabel;
            if (onClickedAction != null) Clicked += onClickedAction;
        }
        internal RibbonButton(string itemId, ResourceManager mgr, bool visible, bool enabled, RdControlSize size,
                IPictureDisp image, bool showImage, bool showLabel, EventHandler onClickedAction)
            : base(itemId, mgr, visible, enabled) {
            _size      = size;
            _image     = new ImageObject(image);
            _showImage = showImage;
            _showLabel = showLabel;
            if (onClickedAction != null) Clicked += onClickedAction;
        }

        #region ISizeableDecoration
        /// <inheritdoc/>
        public RdControlSize Size {
            get { return _size; }
            set { _size = value; OnChanged(); }
        }
        private RdControlSize _size;
        #endregion

        #region IActionableDecoration
        /// <inheritdoc/>
        public event EventHandler Clicked;
        /// <inheritdoc/>
        public void OnAction() => Clicked?.Invoke(this, null);
        #endregion

        #region IImageableDecoration
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
        public void SetImage(IPictureDisp image) { _image = new ImageObject(image);    OnChanged(); }
        /// <inheritdoc/>
        public void SetImageMso(string imageMso) { _image = new ImageObject(imageMso); OnChanged(); }
        #endregion
    }
}
