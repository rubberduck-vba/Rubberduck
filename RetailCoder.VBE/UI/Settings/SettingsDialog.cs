using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Rubberduck.Config;

namespace Rubberduck.UI.Settings
{
    public partial class SettingsDialog : Form
    {
        private Configuration _config;
        private ConfigurationTreeViewControl _treeview;
        private Control _settingsControl;

        public SettingsDialog()
        {
            InitializeComponent();

            _config = ConfigurationLoader.LoadConfiguration();
            _treeview = new ConfigurationTreeViewControl(_config);

            var markers = new List<ToDoMarker>(_config.UserSettings.ToDoListSettings.ToDoMarkers);
            _settingsControl = new TodoListSettingsControl(new TodoSettingModel(markers));
           
            this.splitContainer1.Panel1.Controls.Add(_treeview);
            this.splitContainer1.Panel2.Controls.Add(_settingsControl);

            _treeview.Dock = DockStyle.Fill;
            _settingsControl.Dock = DockStyle.Fill;

            RegisterEvents();   
        }

        private void RegisterEvents()
        {
            _treeview.NodeSelected += _treeview_NodeSelected;
           
        }

        private void _treeview_NodeSelected(object sender, TreeViewEventArgs e)
        {
            if (e.Node.Text == "Rubberduck")
            {
                return;
            }

            if (e.Node.Text == "Todo List")
            {
                //todo: activate todolist control in Panel2
                return;
            }

            if (e.Node.Text == "Code Inpsections")
            {
                //todo: activate inspection control in Panel2
                return;
            }
            throw new NotImplementedException();
        }


    }
}
