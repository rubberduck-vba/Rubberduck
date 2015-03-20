using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Rubberduck.ToDoItems;

namespace Rubberduck.UI.ToDoItems
{
    public interface IToDoExplorerWindow : IDockableUserControl
    {
        DataGridView GridView { get; }
        event EventHandler<ToDoItemClickEventArgs> NavigateToDoItem;
        event EventHandler RefreshToDoItems;
        event EventHandler<DataGridViewCellMouseEventArgs> SortColumn;
        bool SortedAscending { get; set; }
        string SortedByColumn { get; set; }
        IEnumerable<ToDoItem> TodoItems { get; set; }
    }
}
