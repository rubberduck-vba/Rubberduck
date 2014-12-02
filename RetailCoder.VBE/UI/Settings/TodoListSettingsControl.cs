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
        private TodoSettingView _view;
        private IToDoMarker _activeMarker;

        /// <summary>   Parameterless Constructor is to enable design view only. DO NOT USE. </summary>
        public TodoListSettingsControl()
        {
            InitializeComponent();
        }

        public TodoListSettingsControl(TodoSettingView view):this()
        {
            _view = view;
            this.tokenListBox.DataSource = _view.Markers;
            this.tokenListBox.SelectedIndex = 0;
            this.priorityComboBox.DataSource = Enum.GetValues(typeof(Config.TodoPriority));

            //todo: disable combo and text box until edit button is clicked

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

    }
}
