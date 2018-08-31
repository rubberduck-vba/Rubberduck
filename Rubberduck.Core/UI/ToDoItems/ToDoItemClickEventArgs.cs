using System;
using Rubberduck.ToDoItems;

namespace Rubberduck.UI.ToDoItems
{
    public class ToDoItemClickEventArgs : EventArgs
    {
        public ToDoItemClickEventArgs(ToDoItem selectedItem)
        {
            SelectedItem = selectedItem;
        }

        public ToDoItem SelectedItem { get; }
    }
}
