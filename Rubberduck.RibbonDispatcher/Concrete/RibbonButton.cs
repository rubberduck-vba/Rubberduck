using System;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;

using Rubberduck.RibbonDispatcher.Abstract;

namespace Rubberduck.RibbonDispatcher.Concrete
{
    using LanguageStrings     = IRibbonTextLanguageControl;

    /// <summary>TODO</summary>
    [Serializable]
    [CLSCompliant(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IClickedEvents))]
    public class RibbonButton : RibbonCommon, IRibbonButton {
        internal RibbonButton(string id, LanguageStrings strings, bool visible, bool enabled, MyRibbonControlSize size,
                bool showImage, bool showLabel, EventHandler onClickedAction)
            : base(id, strings, visible, enabled, size, showImage, showLabel){
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
