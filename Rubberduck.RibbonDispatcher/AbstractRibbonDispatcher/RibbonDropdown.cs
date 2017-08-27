using System;
using Microsoft.Office.Core;

namespace RubberDuck.RibbonDispatcher {
    using SelectionMadeEventHandler = EventHandler<SelectionMadeEventArgs>;
    public class RibbonDropDown : RibbonCommon, IRibbonDropDown {
        internal RibbonDropDown(string id, bool visible, bool enabled, RibbonControlSize size)
            : base(id, visible, enabled, size){
        }

        public event SelectionMadeEventHandler Clicked;

        public string SelectedItemId { get; set; }

        public void OnAction(string itemId) => Clicked?.Invoke(this, new SelectionMadeEventArgs(itemId));

        public IRibbonCommon AsRibbonControl => this;
    }
}
