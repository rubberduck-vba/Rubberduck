////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Resources;
using System.Runtime.InteropServices;

using Rubberduck.RibbonDispatcher.Abstract;
using stdole;

namespace Rubberduck.RibbonDispatcher.Concrete {
    /// <summary>TODO</summary>
    [Serializable]
    [CLSCompliant(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IClickedEvents))]
    public class RibbonButton : RibbonCommon, IRibbonButton {
        internal RibbonButton(string id, ResourceManager mgr, bool visible, bool enabled, MyRibbonControlSize size,
                string imageMso, bool showImage, bool showLabel, EventHandler onClickedAction)
            : base(id, mgr, visible, enabled, size){
            _image     = new ImageObject(imageMso);
            _showImage = showImage;
            _showLabel = showLabel;
            if (onClickedAction != null) Clicked += onClickedAction;
        }
        internal RibbonButton(string id, ResourceManager mgr, bool visible, bool enabled, MyRibbonControlSize size,
                IPictureDisp image, bool showImage, bool showLabel, EventHandler onClickedAction)
            : base(id, mgr, visible, enabled, size){
            _image     = new ImageObject(image);
            _showImage = showImage;
            _showLabel = showLabel;
            if (onClickedAction != null) Clicked += onClickedAction;
        }

        /// <inheritdoc/>
        public event EventHandler Clicked;
        /// <inheritdoc/>
        public void OnAction() => Clicked?.Invoke(this, null);

        #region IRibbonImageable implementation
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
