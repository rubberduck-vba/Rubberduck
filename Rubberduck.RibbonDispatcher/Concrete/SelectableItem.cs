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
        public SelectableItem(string ItemId, ResourceManager ResourceManager, IPictureDisp Image) 
            : base(ItemId, ResourceManager, true, true)
            => _image = new ImageObject(Image);
        /// <summary>TODO</summary>
        public SelectableItem(string ItemId, ResourceManager ResourceManager, string ImageMso)
            : base(ItemId, ResourceManager, true, true)
            => _image = new ImageObject(ImageMso);

        #region IImageableItem implementation
        /// <inheritdoc/>
        public object Image => _image.Image;
        private ImageObject _image;
        /// <inheritdoc/>
        public bool ShowLabel {
            get => _showLabel;
            set { _showLabel = value; OnChanged(); }
        }
        private bool _showLabel;
        /// <inheritdoc/>
        public bool ShowImage {
            get => _showImage && Image != null;
            set { _showImage = value; OnChanged(); }
        }
        private bool _showImage;

        /// <inheritdoc/>
        public void SetImage(IPictureDisp Image) { _image = new ImageObject(Image);    OnChanged(); }
        /// <inheritdoc/>
        public void SetImageMso(string ImageMso) { _image = new ImageObject(ImageMso); OnChanged(); }
        #endregion
    }
}
