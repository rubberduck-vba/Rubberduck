using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Settings;
using Rubberduck.ToDoItems;

namespace Rubberduck.UI.Settings
{
    public partial class TodoListSettingsUserControl : UserControl, ITodoSettingsView
    {
        /// <summary>   Parameterless Constructor is to enable design view only. DO NOT USE. </summary>
        public TodoListSettingsUserControl()
        {
            InitializeComponent();
        }

        public TodoListSettingsUserControl(IList<ToDoMarker> markers)
            : this()
        {
            InitTodoMarkersGridView(markers);
        }

        private void InitTodoMarkersGridView(IList<ToDoMarker> markers)
        {
            TodoMarkersGridView.AutoGenerateColumns = false;
            TodoMarkersGridView.Columns.Clear();
            TodoMarkersGridView.DataSource = new BindingList<ToDoMarker>(markers);
            TodoMarkersGridView.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;
            TodoMarkersGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            TodoMarkersGridView.CellValueChanged += SelectedPriorityChanged;

            var markerTextColumn = new DataGridViewTextBoxColumn
            {
                Name = "Text",
                DataPropertyName = "Text",
                HeaderText = "Text",
                ReadOnly = true
            };

            var markerPriorityColumn = new DataGridViewComboBoxColumn
            {
                Name = "Priority",
                DataSource = TodoLabels(),
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                HeaderText = "Priority",
                DataPropertyName = "PriorityLabel",
            };

            TodoMarkersGridView.Columns.AddRange(markerTextColumn, markerPriorityColumn);
        }

        private List<string> TodoLabels()
        {
            return (from object priority in Enum.GetValues(typeof(TodoPriority))
                    select
                    RubberduckUI.ResourceManager.GetString("ToDoPriority_" + priority, RubberduckUI.Culture))
                    .ToList();
        }

        public int SelectedIndex
        {
            get { return TodoMarkersGridView.SelectedRows[0].Index; }
            set { TodoMarkersGridView.Rows[value].Selected = true; }
        }

        public TodoPriority ActiveMarkerPriority
        {
            get { return TodoMarkers[SelectedIndex].Priority; }
            set { TodoMarkersGridView.SelectedRows[0].Cells[1].Value = (int)value; }
        }

        public string ActiveMarkerText 
        {
            get { return TodoMarkers[SelectedIndex].Text; }
            set { TodoMarkersGridView.SelectedRows[0].Cells[0].Value = value; }
        }

        public BindingList<ToDoMarker> TodoMarkers
        {
            get { return (BindingList<ToDoMarker>)TodoMarkersGridView.DataSource; }
            set { TodoMarkersGridView.DataSource = value; }
        }

        public event EventHandler PriorityChanged;
        private void SelectedPriorityChanged(object sender, DataGridViewCellEventArgs e)
        {
            RaiseEvent(this, e, PriorityChanged);
        }

        public event EventHandler AddMarker;
        private void addButton_Click(object sender, EventArgs e)
        {
            RaiseEvent(this, e, AddMarker);
        }

        public event EventHandler RemoveMarker;
        private void removeButton_Click(object sender, EventArgs e)
        {
            RaiseEvent(this, e, RemoveMarker);
        }

        private void RaiseEvent(object sender, EventArgs e, EventHandler handler)
        {
            if (handler != null)
            {
                handler(this, e);
            }
        }
    }
}
