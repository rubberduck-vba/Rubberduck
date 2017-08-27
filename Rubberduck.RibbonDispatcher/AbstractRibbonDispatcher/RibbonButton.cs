using System;
using Microsoft.Office.Core;

namespace Rubberduck.RibbonDispatcher.Abstract {
    using LanguageStrings     = IRibbonTextLanguageControl;

    public class RibbonButton : RibbonCommon, IRibbonButton {
        internal RibbonButton(string id, LanguageStrings strings, bool visible, bool enabled, RibbonControlSize size)
            : base(id, strings, visible, enabled, size){
        }

        public event EventHandler Clicked;

        public bool ShowLabel { get; set; }
        public bool ShowImage { get; set; }

        public void OnAction() => Clicked?.Invoke(this,null);
 
        public IRibbonCommon AsRibbonControl => this;
  }
}
