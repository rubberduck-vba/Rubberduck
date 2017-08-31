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
    public class RibbonButton : RibbonCommon, IRibbonButton {
        internal RibbonButton(string id, ResourceManager mgr, bool visible, bool enabled, MyRibbonControlSize size,
                bool showImage, bool showLabel, EventHandler onClickedAction)
            : base(id, mgr, visible, enabled, size, showImage, showLabel){
            if (onClickedAction != null) Clicked += onClickedAction;
        }

        /// <inheritdoc/>
        public event EventHandler Clicked;

        /// <inheritdoc/>
        public void OnAction() => Clicked?.Invoke(this,null);

        /// <inheritdoc/>
        public IRibbonCommon AsRibbonControl => this;
  }
}
