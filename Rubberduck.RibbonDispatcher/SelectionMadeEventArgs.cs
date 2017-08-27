using System;

namespace Rubberduck.RibbonDispatcher
{
    public class SelectionMadeEventArgs : EventArgs {
        public SelectionMadeEventArgs(string itemId) { ItemId = itemId; }
        public string ItemId    { get; }
    }
}
