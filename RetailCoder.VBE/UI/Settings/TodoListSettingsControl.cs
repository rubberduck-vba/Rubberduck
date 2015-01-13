using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Rubberduck.Config;

namespace Rubberduck.UI.Settings
{
    public partial class TodoListSettingsControl : UserControl, ITodoSettingsView
    {
        private BindingList<ToDoMarker> _markers;
        private ToDoMarker _activeMarker;

        /// <summary>   Parameterless Constructor is to enable design view only. DO NOT USE. </summary>
        public TodoListSettingsControl()
        {
            InitializeComponent();
        }

        public TodoListSettingsControl(List<ToDoMarker> markers)
            : this()
        {
            _markers = new BindingList<ToDoMarker>(markers.ToList());
            this.tokenListBox.DataSource = _markers;
            this.tokenListBox.SelectedIndex = 0;
            this.priorityComboBox.DataSource = Enum.GetValues(typeof(Config.TodoPriority));

            SetActiveMarker();
        }

        private void SetActiveMarker()
        {
            _activeMarker = (ToDoMarker)this.tokenListBox.SelectedItem;
            if (_activeMarker != null && this.priorityComboBox.Items.Count > 0)
            {
                this.priorityComboBox.SelectedIndex = (int)_activeMarker.Priority;
            }

            this.tokenTextBox.Text = _activeMarker.Text;
        }

        private void SaveActiveMarker()
        {
            if (_activeMarker != null && this.priorityComboBox.Items.Count > 0)
            {
                _markers[this.tokenListBox.SelectedIndex] = _activeMarker;
            }
        }

        private void tokenListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetActiveMarker();
        }

        private void saveChangesButton_Click(object sender, EventArgs e)
        {
            var index = this.tokenListBox.SelectedIndex;
            _markers[index].Text = tokenTextBox.Text;
            _markers[index].Priority = (TodoPriority)priorityComboBox.SelectedIndex;
            SaveActiveMarker(); //does this really need to happen? Changes still aren't being serialized.
        }

        private void tokenTextBox_TextChanged(object sender, EventArgs e)
        {
            this.saveChangesButton.Enabled = true;
        }

        private void addButton_Click(object sender, EventArgs e)
        {
            var marker = new ToDoMarker(this.tokenTextBox.Text, (TodoPriority)this.priorityComboBox.SelectedIndex);
            _markers.Add(marker);

            this.tokenListBox.DataSource = _markers;

            //todo: adding an item should shift the selected index of the listbox to the newly added item
        }

        private void removeButton_Click(object sender, EventArgs e)
        {
            _markers.RemoveAt(this.tokenListBox.SelectedIndex);

            this.tokenListBox.DataSource = _markers;
        }

        //interface implementation

        public int SelectedIndex
        {
            get { return this.tokenListBox.SelectedIndex; }
            set { this.tokenListBox.SelectedIndex = value; }
        }

        public TodoPriority ActiveMarkerPriority
        {
            get
            {
                throw new NotImplementedException();
            }
            set
            {
                throw new NotImplementedException();
            }
        }

        public string ActiveMarkerText
        {
            get
            {
                throw new NotImplementedException();
            }
            set
            {
                throw new NotImplementedException();
            }
        }

        public BindingList<ToDoMarker> TodoMarkers
        {
            get
            {
                throw new NotImplementedException();
            }
            set
            {
                throw new NotImplementedException();
            }
        }

        public event EventHandler RemoveMarker;

        public event EventHandler AddMarker;

        public event EventHandler SaveMarker;

        public event EventHandler SelectionChanged;
    }
}
