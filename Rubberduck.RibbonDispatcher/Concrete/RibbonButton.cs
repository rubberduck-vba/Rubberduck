using System;
using System.Runtime.InteropServices;
using System.Diagnostics.CodeAnalysis;
using stdole;

using Rubberduck.RibbonDispatcher.ControlMixins;
using Rubberduck.RibbonDispatcher.AbstractCOM;

namespace Rubberduck.RibbonDispatcher.Concrete {
    /// <summary>The ViewModel for Ribbon Button objects.</summary>
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable",
        Justification = "Publc, Non-Creatable, class with exported Events.")]
    [Serializable]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IClickedEvents))]
    [ComDefaultInterface(typeof(IRibbonButton))]
    [Guid(RubberduckGuid.RibbonButton)]
    public class RibbonButton : RibbonCommon, IRibbonButton,
        ISizeableMixin, IClickableMixin, IImageableMixin {
        internal RibbonButton(string itemId, IResourceManager mgr, bool visible, bool enabled, RdControlSize size,
                ImageObject image, bool showImage, bool showLabel) : base(itemId, mgr, visible, enabled) {
            this.SetSize(size);
            this.SetImage(image);
            this.SetShowImage(showImage);
            this.SetShowLabel(showLabel);
        }

        #region Publish ISizeableMixin to class default interface
        /// <summary>Gets or sets the preferred {RdControlSize} for the control.</summary>
        public RdControlSize Size {
            get => this.GetSize();
            set => this.SetSize(value);
        }
        #endregion

        #region Publish IActionableDecoration to class default interface
        /// <summary>The Clicked event source for COM clients</summary>
        public event ClickedEventHandler Clicked;

        /// <summary>The callback from the Ribbon Dispatcher to initiate Clicked events on this control.</summary>
        public void OnClicked() => Clicked?.Invoke();
        #endregion

        #region Publish IImageableMixin to class default interface
        /// <inheritdoc/>
        public object Image => this.GetImage();
        /// <summary>Gets or sets whether the image for this control should be displayed when its size is {rdRegular}.</summary>
        public bool ShowImage {
            get => this.GetShowImage();
            set => this.SetShowImage(value);
        }
        /// <summary>Gets or sets whether the label for this control should be displayed when its size is {rdRegular}.</summary>
        public bool ShowLabel {
            get => this.GetShowLabel();
            set => this.SetShowLabel(value);
        }

        /// <summary>Sets the displayable image for this control to the provided {IPictureDisp}</summary>
        public void SetImageDisp(IPictureDisp Image) => this.SetImage(Image);
        /// <summary>Sets the displayable image for this control to the named ImageMso image</summary>
        public void SetImageMso(string ImageMso)     => this.SetImage(ImageMso);
        #endregion
    }
}
