using System;
using System.Runtime.InteropServices;
using stdole;

using Rubberduck.RibbonDispatcher.AbstractCOM;
using Rubberduck.RibbonDispatcher.ControlMixins;

namespace Rubberduck.RibbonDispatcher.Concrete {
    /// <summary>TODO</summary>
    [Serializable]
    [CLSCompliant(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ISelectableItem))]
    public class SelectableItem : RibbonCommon, ISelectableItem, IImageableMixin {
        /// <summary>TODO</summary>
        internal SelectableItem(string ItemId, IResourceManager ResourceManager, ImageObject Image) 
            : base(ItemId, ResourceManager, true, true)
            => this.SetImage(Image, null);

        #region Publish IImageableMixin to class default interface
        /// <inheritdoc/>
        public object Image => this.GetImage();
        /// <summary>Gets or sets whether the image for this control should be displayed when its size is {rdRegular}.</summary>
        public bool ShowImage {
            get => this.GetShowImage();
            set => this.SetShowImage(value, OnChanged);
        }
        /// <summary>Gets or sets whether the label for this control should be displayed when its size is {rdRegular}.</summary>
        public bool ShowLabel {
            get => this.GetShowLabel();
            set => this.SetShowLabel(value, OnChanged);
        }

        /// <summary>Sets the displayable image for this control to the provided {IPictureDisp}</summary>
        public void SetImageDisp(IPictureDisp Image) => this.SetImage(Image, OnChanged);
        /// <summary>Sets the displayable image for this control to the named ImageMso image</summary>
        public void SetImageMso(string ImageMso)     => this.SetImage(ImageMso, OnChanged);
        #endregion
    }
}
