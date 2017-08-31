using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using stdole;

using Rubberduck.RibbonDispatcher.ControlMixins;
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
    public class RibbonToggleButton : RibbonCommon, IRibbonToggleButton,
        ISizeableMixin, IToggleableMixin, IImageableMixin {
        internal RibbonToggleButton(string itemId, IResourceManager mgr, bool visible, bool enabled, RdControlSize size,
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

        #region Publish IToggleableMixin to class default interface
        /// <summary>TODO</summary>
        public event ToggledEventHandler Toggled;

        /// <summary>TODO</summary>
        public          bool   IsPressed => this.GetPressed();
        /// <summary>TODO</summary>
        public override string Label     => this.GetLabel();

        /// <summary>TODO</summary>
        public void OnActionToggle(bool IsPressed) => this.OnActionToggled(IsPressed, b => Toggled?.Invoke(b));
        /// <summary>TODO</summary>
        IRibbonTextLanguageControl IToggleableMixin.LanguageStrings => LanguageStrings;
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
