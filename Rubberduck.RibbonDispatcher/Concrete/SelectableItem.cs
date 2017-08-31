using System;
using System.Resources;
using System.Runtime.InteropServices;
using stdole;

using Rubberduck.RibbonDispatcher.AbstractCOM;

namespace Rubberduck.RibbonDispatcher.Concrete {
    /// <summary>TODO</summary>
    [Serializable]
    [CLSCompliant(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ISelectableItem))]
    public class SelectableItem : RibbonCommon, ISelectableItem {
        /// <summary>TODO</summary>
        public SelectableItem(string itemId, ResourceManager resourceManager, IPictureDisp image) 
            : base(itemId, resourceManager, true, true) {
            _image = new ImageObject(image);
        }
        /// <summary>TODO</summary>
        public SelectableItem(string itemId, ResourceManager resourceManager, string imageMso)
            : base(itemId, resourceManager, true, true) {
            _image = new ImageObject(imageMso);
        }

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
        public void SetImage(IPictureDisp image) { _image = new ImageObject(image);    OnChanged(); }
        /// <inheritdoc/>
        public void SetImageMso(string imageMso) { _image = new ImageObject(imageMso); OnChanged(); }
        #endregion
    }
}
