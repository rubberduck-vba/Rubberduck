using System;

namespace Rubberduck.UI.IdentifierReferences
{
    public class ListItemActionEventArgs : EventArgs
    {
        public ListItemActionEventArgs(object selectedItem)
        {
            _selectedItem = selectedItem;
        }

        private readonly object _selectedItem;
        public object SelectedItem { get { return _selectedItem; } }
    }
}
