using System;

namespace Rubberduck.RibbonDispatcher
{
    [CLSCompliant(true)]
    public class SelectionMadeEventArgs : EventArgs {
        public SelectionMadeEventArgs(string itemId) { ItemId = itemId; }
        public string ItemId    { get; }
    }
}
