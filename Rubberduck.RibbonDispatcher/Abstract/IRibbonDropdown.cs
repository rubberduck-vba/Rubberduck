using System;
using System.Runtime.InteropServices;

using Rubberduck.RibbonDispatcher;

namespace Rubberduck.RibbonDispatcher.Abstract
{
    using SelectionMadeEventHandler = EventHandler<SelectionMadeEventArgs>;

    [ComVisible(true)][CLSCompliant(true)]
    public interface IRibbonDropDown {
        event SelectionMadeEventHandler Clicked;

        string SelectedItemId { get; set; }

        void OnAction(string itemId);

        IRibbonCommon AsRibbonControl { get; }
    }
}
