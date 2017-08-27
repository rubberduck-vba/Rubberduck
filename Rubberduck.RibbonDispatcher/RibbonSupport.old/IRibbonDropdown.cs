using System;
using System.Runtime.InteropServices;

namespace RubberDuck.RibbonSupport {
    using SelectionMadeEventHandler = EventHandler<SelectionMadeEventArgs>;

    [ComVisible(true)][CLSCompliant(true)]
    public interface IRibbonDropdown {
        event SelectionMadeEventHandler Clicked;

        string SelectedItemID { get; set; }

        void OnAction(string itemID);

        IRibbonCommon AsRibbonControl { get; }
    }
}
