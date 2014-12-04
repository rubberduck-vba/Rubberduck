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
    public partial class TodoListSettingsControl : UserControl
    {
        private TodoSettingModel _model;
        private IToDoMarker _activeMarker;

        /// <summary>   Parameterless Constructor is to enable design view only. DO NOT USE. </summary>
        public TodoListSettingsControl()
        {
            InitializeComponent();
        }

        public TodoListSettingsControl(TodoSettingModel model):this()
        {
            _model = model;
            this.tokenListBox.DataSource = _model.Markers;
            this.tokenListBox.SelectedIndex = 0;
            this.priorityComboBox.DataSource = Enum.GetValues(typeof(Config.TodoPriority));

            SetActiveMarker();
        }

        private void SetActiveMarker()
        {
            _activeMarker = (IToDoMarker)this.tokenListBox.SelectedItem;
            if (_activeMarker != null && this.priorityComboBox.Items.Count > 0)
            {
                this.priorityComboBox.SelectedIndex = _activeMarker.Priority;
            }

            this.tokenTextBox.Text = _activeMarker.Text;
        }

        private void tokenListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetActiveMarker();
        }

        private void saveChangesButton_Click(object sender, EventArgs e)
        {
            var index = this.tokenListBox.SelectedIndex;
            _model.Markers[index].Text = tokenTextBox.Text;
            _model.Markers[index].Priority = priorityComboBox.SelectedIndex;
            _model.Save();
        }

        private void tokenTextBox_TextChanged(object sender, EventArgs e)
        {
            this.saveChangesButton.Enabled = true;
        }

        private void addButton_Click(object sender, EventArgs e)
        {
            var marker = new ToDoMarker(this.tokenTextBox.Text, this.priorityComboBox.SelectedIndex);
            _model.Markers.Add(marker);
            _model.Save();

            this.tokenListBox.DataSource = _model.Markers;
        }

        private void removeButton_Click(object sender, EventArgs e)
        {
            _model.Markers.RemoveAt(this.tokenListBox.SelectedIndex);
            _model.Save();

            this.tokenListBox.DataSource = _model.Markers;
        }

    }
}
