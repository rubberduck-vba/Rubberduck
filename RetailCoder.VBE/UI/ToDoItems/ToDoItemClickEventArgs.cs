using System;
using Rubberduck.ToDoItems;

namespace Rubberduck.UI.ToDoItems
{
    public class ToDoItemClickEventArgs : EventArgs
    {
        public ToDoItemClickEventArgs(ToDoItem selectedItem)
        {
            _selectedItem = selectedItem;
        }

        private readonly ToDoItem _selectedItem;
        public ToDoItem SelectedItem { get { return _selectedItem; } }
    }
}
