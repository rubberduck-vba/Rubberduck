using System;
using System.Resources;
using System.Runtime.InteropServices;

using Rubberduck.RibbonDispatcher.Abstract;

namespace Rubberduck.RibbonDispatcher.Concrete {
    /// <summary>TODO</summary>
    [Serializable]
    [CLSCompliant(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IClickedEvents))]
    public class RibbonButton : RibbonCommon, IRibbonButton, IRibbonImageable {
        internal RibbonButton(string id, ResourceManager mgr, bool visible, bool enabled, MyRibbonControlSize size,
                bool showImage, bool showLabel, EventHandler onClickedAction)
            : base(id, mgr, visible, enabled, size){
            _showImage = showImage;
            _showLabel = showLabel;
            if (onClickedAction != null) Clicked += onClickedAction;
        }

        /// <inheritdoc/>
        public bool ShowLabel {
            get { return _showLabel; }
            set { _showLabel = value; OnChanged(); }
        }
        private bool _showLabel;
        /// <inheritdoc/>
        public bool ShowImage {
            get { return _showImage; }
            set { _showImage = value; OnChanged(); }
        }
        private bool _showImage;

        /// <inheritdoc/>
        public event EventHandler Clicked;

        /// <inheritdoc/>
        public void OnAction() => Clicked?.Invoke(this,null);
  }
}
