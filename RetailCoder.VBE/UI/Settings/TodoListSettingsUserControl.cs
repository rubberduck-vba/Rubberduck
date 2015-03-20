using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Forms;
using Rubberduck.Config;

namespace Rubberduck.UI.Settings
{
    public partial class TodoListSettingsUserControl : UserControl, ITodoSettingsView
    {
        /// <summary>   Parameterless Constructor is to enable design view only. DO NOT USE. </summary>
        public TodoListSettingsUserControl()
        {
            InitializeComponent();
        }

        public TodoListSettingsUserControl(List<ToDoMarker> markers)
            : this()
        {
            this.tokenListBox.DataSource = new BindingList<ToDoMarker>(markers);
            this.tokenListBox.SelectedIndex = 0;
            this.priorityComboBox.DataSource = Enum.GetValues(typeof(TodoPriority));
        }

        public int SelectedIndex
        {
            get { return this.tokenListBox.SelectedIndex; }
            set { this.tokenListBox.SelectedIndex = value; }
        }

        public bool SaveEnabled
        {
            get { return this.saveChangesButton.Enabled; }
            set { this.saveChangesButton.Enabled = value; }
        }

        public TodoPriority ActiveMarkerPriority
        {
            get { return (TodoPriority)this.priorityComboBox.SelectedIndex; }
            set { this.priorityComboBox.SelectedIndex = (int)value; }
        }

        public string ActiveMarkerText 
        {
            get { return this.tokenTextBox.Text; }
            set { this.tokenTextBox.Text = value; }
        }

        public BindingList<ToDoMarker> TodoMarkers
        {
            get { return (BindingList<ToDoMarker>)this.tokenListBox.DataSource; }
            set { this.tokenListBox.DataSource = value; }
        }

        public event EventHandler SelectionChanged;
        private void tokenListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            RaiseEvent(this, e, SelectionChanged);
        }

        public event EventHandler SaveMarker;
        private void saveChangesButton_Click(object sender, EventArgs e)
        {
            RaiseEvent(this, e, SaveMarker);
        }

        public event EventHandler TextChanged;
        private void tokenTextBox_TextChanged(object sender, EventArgs e)
        {
            RaiseEvent(this, e, TextChanged);
        }

        public event EventHandler PriorityChanged;
        private void priorityComboBox_SelectedIndexChanged(object sender, EventArgs e)
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
