using System;

namespace Rubberduck.UI.IdentifierReferences
{
    public class ListItemActionEventArgs : EventArgs
    {
        public ListItemActionEventArgs(object selectedItem)
        {
            SelectedItem = selectedItem;
        }

        public object SelectedItem { get; }
    }
}
