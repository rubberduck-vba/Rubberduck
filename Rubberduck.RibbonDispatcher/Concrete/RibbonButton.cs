using System;
using System.Runtime.InteropServices;
using System.Diagnostics.CodeAnalysis;
using stdole;

using Rubberduck.RibbonDispatcher.ControlDecorators;
using Rubberduck.RibbonDispatcher.AbstractCOM;
using Rubberduck.RibbonDispatcher.EventHandlers;

namespace Rubberduck.RibbonDispatcher.Concrete {
    /// <summary>The ViewModel for Ribbon Button objects.</summary>
    [Serializable]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IClickedComEvents))]
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable",
        Justification = "Publc, Non-Creatable class with exported Events.")]
    [ComDefaultInterface(typeof(IRibbonButton))]
    [Guid(RubberduckGuid.RibbonButton)]
    public class RibbonButton : RibbonCommon, IRibbonButton,
        ISizeableMixin, IActionableDecorator, IImageableDecorator {
        internal RibbonButton(string itemId, IResourceManager mgr, bool visible, bool enabled, RdControlSize size,
                string imageMso, bool showImage, bool showLabel)
            : this(itemId, mgr, visible, enabled, size, new ImageObject(imageMso), showImage, showLabel) { }
        internal RibbonButton(string itemId, IResourceManager mgr, bool visible, bool enabled, RdControlSize size,
                IPictureDisp image, bool showImage, bool showLabel)
            : this(itemId, mgr, visible, enabled, size, new ImageObject(image), showImage, showLabel) { }
        private RibbonButton(string itemId, IResourceManager mgr, bool visible, bool enabled, RdControlSize size,
                ImageObject image, bool showImage, bool showLabel) : base(itemId, mgr, visible, enabled) {
            this.SetSize(size, null);
            _image     = image;
            _showImage = showImage;
            _showLabel = showLabel;
        }

        #region ISizeableMixin
        /// <summary>Sets ofr gets the preferred {RdControlSize} for the control.</summary>
        public RdControlSize Size {
            get => this.GetSize();
            set => this.SetSize(value, OnChanged);
        }
        #endregion

        #region IActionableDecoration
        /// <summary>The Clicked event source for DOT NET clients</summary>
        public event EventHandler<EventArgs> Clicked;
        /// <summary>The Clicked event source for COM clients</summary>
        public event ClickedComEventHandler ComClicked;
        /// <summary>The callback from the Ribbon Dispatcher to initiate Clicked events on this control.</summary>
        public void OnAction() {
            Clicked?.Invoke(this, EventArgs.Empty);
            ComClicked?.Invoke();
        }
        #endregion

        #region IImageableDecoration
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
