////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Resources;
using System.Runtime.InteropServices;
using stdole;

using Rubberduck.RibbonDispatcher.AbstractCOM;
using Rubberduck.RibbonDispatcher.EventHandlers;

namespace Rubberduck.RibbonDispatcher.Concrete {
    /// <summary>Returns a new Ribbon DropDownViewModel instance.</summary>
    [Serializable]
    [CLSCompliant(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(ISelectionMadeEvents))]
    [ComDefaultInterface(typeof(IRibbonDropDown))]
    [Guid(RubberduckGuid.RibbonDropDown)]
    public class RibbonDropDown : RibbonCommon, IRibbonDropDown {
        internal RibbonDropDown(string id, ResourceManager mgr, bool visible, bool enabled, RdControlSize size)
            : base(id, mgr, visible, enabled, size){
        }

        /// <summary>TODO</summary>
        public event SelectionMadeEventHandler SelectionMade;

        /// <summary>TODO</summary>
        public string SelectedItemId { get; set; }

        /// <summary>TODO</summary>
        public void OnActionDropDown(string itemId) => SelectionMade?.Invoke(this, new SelectionMadeEventArgs(itemId));

        /// <summary>TODO</summary>
        public IRibbonCommon AsRibbonControl => this;

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
