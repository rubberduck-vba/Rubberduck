using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Rubberduck.ToDoItems;

namespace Rubberduck.UI.ToDoItems
{
    [ComVisible(false)]
    public partial class ToDoExplorerWindow : UserControl, IDockableUserControl
    {
        private const string ClassId = "8B071EDA-2C9C-4009-9A22-A1958BF98B28";
        string IDockableUserControl.ClassId { get { return ClassId; } }
        string IDockableUserControl.Caption { get { return "ToDo Explorer"; } }

        private readonly BindingList<ToDoItem> _todoItems;

        public ToDoExplorerWindow()
            : this(new ToDoItem[]{})
        { }

        public ToDoExplorerWindow(IEnumerable<ToDoItem> items)
        {
            _todoItems = new BindingList<ToDoItem>(items.ToList());
            InitializeComponent();
            InitializeGrid();
        }

        private void InitializeGrid()
        {
            todoItemsGridView.DataSource = _todoItems;

            todoItemsGridView.Columns["Description"].FillWeight = 150;
            todoItemsGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            
            todoItemsGridView.CellDoubleClick += ToDoGridViewCellDoubleClicked;
            refreshButton.Click += RefreshButtonClicked;
        }

        public event EventHandler<ToDoItemClickEventArgs> NavigateToDoItem;
        private void ToDoGridViewCellDoubleClicked(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0)
            {
                return;
            }

            var handler = NavigateToDoItem;
            if (handler == null)
            {
                return;
            }

            var item = (ToDoItem)todoItemsGridView[e.ColumnIndex, e.RowIndex].OwningRow.DataBoundItem;
            var args = new ToDoItemClickEventArgs(item);
            handler(this, args);
        }

        public event EventHandler RefreshToDoItems;
        private void RefreshButtonClicked(object sender, EventArgs e)
        {
            var handler = RefreshToDoItems;
            if (handler == null)
            {
                return;
            }

            handler(this, EventArgs.Empty);
        }

        public void SetItems(IEnumerable<ToDoItem> items)
        {
            _todoItems.Clear();
            foreach (var toDoItem in items)
            {
                _todoItems.Add(toDoItem);
            }
        }
    }
}
