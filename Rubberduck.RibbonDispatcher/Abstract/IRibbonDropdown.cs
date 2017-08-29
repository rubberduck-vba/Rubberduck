using System;
using System.Runtime.InteropServices;

namespace Rubberduck.RibbonDispatcher.Abstract
{
    using SelectionMadeEventHandler = EventHandler<SelectionMadeEventArgs>;

    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonDropDown : IRibbonCommon
    {
        event  SelectionMadeEventHandler Clicked;

        string SelectedItemId  { get; set; }

        void   OnAction(string itemId);
    }
}
