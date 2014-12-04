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

        public TodoListSettingsControl(TodoSettingModel view):this()
        {
            _model = view;
            this.tokenListBox.DataSource = _model.Markers;
            this.tokenListBox.SelectedIndex = 0;
            this.priorityComboBox.DataSource = Enum.GetValues(typeof(Config.TodoPriority));

            SetActiveMarker();
        }

        private void TodoListSettingsControl_Load(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void priorityComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void SetActiveMarker()
        {
            _activeMarker = (IToDoMarker)this.tokenListBox.SelectedItem;
            if (this.priorityComboBox.Items.Count > 0)
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

    }
}
