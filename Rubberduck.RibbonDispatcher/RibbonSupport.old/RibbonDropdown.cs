using System;
using Microsoft.Office.Core;

namespace RubberDuck.RibbonSupport {
    using SelectionMadeEventHandler = EventHandler<SelectionMadeEventArgs>;
    public class RibbonDropdown : RibbonCommon, IRibbonDropdown {
        internal RibbonDropdown(string id, bool visible, bool enabled, RibbonControlSize size)
            : base(id, visible, enabled, size){
        }

        public event SelectionMadeEventHandler Clicked;

        public string SelectedItemID { get; set; }

        public void OnAction(string itemID) => Clicked?.Invoke(this, new SelectionMadeEventArgs(itemID));

        public IRibbonCommon AsRibbonControl => this;
    }
}
