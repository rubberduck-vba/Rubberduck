using System;
using System.Runtime.InteropServices;
using Rubberduck.ToDoItems;

namespace Rubberduck.UI.ToDoItems
{
    [ComVisible(false)]
    public class ToDoItemClickEventArgs : EventArgs
    {
        public ToDoItemClickEventArgs(ToDoItem selection)
        {
            _selection = selection;
        }

        private readonly ToDoItem _selection;
        public ToDoItem Selection { get { return _selection; } }
    }
}