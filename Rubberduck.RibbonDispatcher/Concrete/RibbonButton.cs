using System;

using Microsoft.Office.Core;

using Rubberduck.RibbonDispatcher.Abstract;

namespace Rubberduck.RibbonDispatcher.Concrete
{
    using LanguageStrings     = IRibbonTextLanguageControl;

    public class RibbonButton : RibbonCommon, IRibbonButton {
        internal RibbonButton(string id, LanguageStrings strings, bool visible, bool enabled, RibbonControlSize size,
                bool showImage, bool showLabel, EventHandler onClickedAction)
            : base(id, strings, visible, enabled, size, showImage, showLabel){
            if (onClickedAction != null) Clicked += onClickedAction;
        }

        public event EventHandler Clicked;

        public void OnAction() => Clicked?.Invoke(this,null);
 
        public IRibbonCommon AsRibbonControl => this;
  }
}
