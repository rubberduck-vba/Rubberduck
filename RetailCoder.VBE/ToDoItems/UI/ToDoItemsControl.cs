using System;
using System.ComponentModel;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using System.Collections.Generic;

namespace Rubberduck.ToDoItems
{
    public partial class ToDoItemsControl : UserControl
    {
        private VBE vbe;
        private BindingList<ToDoItem> taskList;
        private List<Config.ToDoMarker> markers;

        public ToDoItemsControl(VBE vbe, List<Config.ToDoMarker> markers)
        {
            this.vbe = vbe;
            this.markers = markers;

            InitializeComponent();
            
            RefreshTaskList();
            InitializeGrid();

        }

        private void InitializeGrid()
        {
            todoItemsGridView.DataSource = taskList;
            var descriptionColumn = todoItemsGridView.Columns["Description"];
            if (descriptionColumn != null)
            {
                descriptionColumn.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            }

            todoItemsGridView.CellDoubleClick += taskListGridView_CellDoubleClick;
            refreshButton.Click += refreshButton_Click;

        }

        void taskListGridView_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            ToDoItem task = taskList.ElementAt(e.RowIndex);
            VBComponent component = vbe.ActiveVBProject.VBComponents.Item(task.Module);

            component.CodeModule.CodePane.Show();
            component.CodeModule.CodePane.SetSelection(task.LineNumber, 1, task.LineNumber, 1);
        }

        private void RefreshGridView()
        {
            RefreshTaskList();
            todoItemsGridView.DataSource = taskList;
            todoItemsGridView.Refresh();
        }

        public void RefreshTaskList()
        {
            this.taskList = new BindingList<ToDoItem>();

            foreach (VBComponent component in this.vbe.ActiveVBProject.VBComponents)
            {
                CodeModule module = component.CodeModule;
                for (var i = 1; i <= module.CountOfLines; i++)
                {
                    string line = module.get_Lines(i, 1);
                    Config.ToDoMarker marker;

                    if (TryGetMarker(line, out marker))
                    {
                        var priority = (TaskPriority)marker.priority;
                        this.taskList.Add(new ToDoItem(priority, line, module, i));
                    }
                }
            }
        }

        private bool TryGetMarker(string line, out Config.ToDoMarker result)
        {
            var upCasedLine = line.ToUpper();
            foreach (var marker in this.markers)
            {
                if (upCasedLine.Contains(marker.text))
                {
                    result = marker;
                    return true;
                }
            }
            result = null;
            return false;
        }

        private void refreshButton_Click(object sender, EventArgs e)
        {
            RefreshGridView();
        }
    }
}
