using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.ToDoItems;

namespace Rubberduck.UI.ToDoItems
{
    [ExcludeFromCodeCoverage]
    public partial class ToDoExplorerWindow : UserControl, IToDoExplorerWindow
    {
        private const string ClassId = "8B071EDA-2C9C-4009-9A22-A1958BF98B28";
        string IDockableUserControl.ClassId { get { return ClassId; } }
        string IDockableUserControl.Caption { get { return RubberduckUI.ToDoExplorer_Caption; } }

        private BindingList<ToDoItem> _todoItems;
        public IEnumerable<ToDoItem> TodoItems 
        { 
            get { return _todoItems; }
            set 
            { 
                _todoItems = new BindingList<ToDoItem>(value.ToList());
                todoItemsGridView.DataSource = _todoItems;
                todoItemsGridView.Refresh();
            }
        }

        public DataGridView GridView { get { return todoItemsGridView; } }

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

            todoItemsGridView.Columns["Priority"].Visible = false;
            todoItemsGridView.Columns["PriorityLabel"].HeaderText = RubberduckUI.Priority;
            todoItemsGridView.Columns["Description"].HeaderText = RubberduckUI.TodoExplorer_Description;
            todoItemsGridView.Columns["ProjectName"].HeaderText = RubberduckUI.ProjectName;
            todoItemsGridView.Columns["ModuleName"].HeaderText = RubberduckUI.ModuleName;
            todoItemsGridView.Columns["LineNumber"].HeaderText = RubberduckUI.TodoExplorer_LineNumber;

            refreshButton.ToolTipText = RubberduckUI.Refresh;
        }

        public event EventHandler<ToDoItemClickEventArgs> NavigateToDoItem;
        private void ToDoGridViewCellDoubleClicked(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0)
            {
                return;
            }

            var handler = NavigateToDoItem;
            if (handler != null)
            {
                var item = (ToDoItem)todoItemsGridView[e.ColumnIndex, e.RowIndex].OwningRow.DataBoundItem;
                var args = new ToDoItemClickEventArgs(item);
                handler(this, args);
            }
        }

        public event EventHandler RefreshToDoItems;
        private void RefreshButtonClicked(object sender, EventArgs e)
        {
            var handler = RefreshToDoItems;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            } 
        }

        public event EventHandler<DataGridViewCellMouseEventArgs> SortColumn;
        private void ColumnHeaderMouseClicked(object sender, DataGridViewCellMouseEventArgs e)
        {
            var handler = SortColumn;
            if (handler != null)
            {
                handler(this, e);
            } 
        }
    }
}
