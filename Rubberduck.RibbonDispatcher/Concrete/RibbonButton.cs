using System;
using System.Runtime.InteropServices;
using System.Diagnostics.CodeAnalysis;
using stdole;

using Rubberduck.RibbonDispatcher.ControlMixins;
using Rubberduck.RibbonDispatcher.AbstractCOM;
using Rubberduck.RibbonDispatcher.EventHandlers;

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
        ISizeableMixin, IActionableMixin, IImageableMixin {
        internal RibbonButton(string itemId, IResourceManager mgr, bool visible, bool enabled, RdControlSize size,
                ImageObject image, bool showImage, bool showLabel) : base(itemId, mgr, visible, enabled) {
            this.SetSize(size, null);
            this.SetImage(image, null);
            this.SetShowImage(showImage, null);
            this.SetShowLabel(showLabel, null);
        }

        #region Publish ISizeableMixin to class default interface
        /// <summary>Gets or sets the preferred {RdControlSize} for the control.</summary>
        public RdControlSize Size {
            get => this.GetSize();
            set => this.SetSize(value, OnChanged);
        }
        #endregion

        #region IActionableDecoration
        /// <summary>The Clicked event source for COM clients</summary>
        public event ClickedEventHandler Clicked;
        /// <summary>The callback from the Ribbon Dispatcher to initiate Clicked events on this control.</summary>
        public void OnAction() => Clicked?.Invoke();
        #endregion

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
