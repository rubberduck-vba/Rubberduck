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

//todo: this class needs serious clean up

namespace Rubberduck.UI.Settings
{
    public partial class SettingsDialog : Form
    {
        private Configuration _config;
        private ConfigurationTreeViewControl _treeview;
        private Control _todoList;
        private Control _inspections;

        public SettingsDialog()
        {
            InitializeComponent();

            _config = ConfigurationLoader.LoadConfiguration();
            _treeview = new ConfigurationTreeViewControl(_config);

            var markers = new List<ToDoMarker>(_config.UserSettings.ToDoListSettings.ToDoMarkers);
            _todoList = new TodoListSettingsControl(new TodoSettingModel(markers));
           
            this.splitContainer1.Panel1.Controls.Add(_treeview);
            this.splitContainer1.Panel2.Controls.Add(_todoList);

            _treeview.Dock = DockStyle.Fill;
            _todoList.Dock = DockStyle.Fill;

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
                if (_todoList == null)
                {
                    var markers = new List<ToDoMarker>(_config.UserSettings.ToDoListSettings.ToDoMarkers);
                    _todoList = new TodoListSettingsControl(new TodoSettingModel(markers));
                    _todoList.Dock = DockStyle.Fill;
                }

                this.splitContainer1.Panel2.Controls.Clear();
                this.splitContainer1.Panel2.Controls.Add(_todoList);
            }

            if (e.Node.Text == "Code Inpsections")
            {
                if (_inspections == null)
                {
                    //note: might want to just pass an enumerable instead
                    _inspections = new CodeInspectionControl(_config);
                    _inspections.Dock = DockStyle.Fill;
                }

                this.splitContainer1.Panel2.Controls.Clear();
                this.splitContainer1.Panel2.Controls.Add(_inspections);
                
            }
        }

        private void SettingsDialog_FormClosed(object sender, FormClosedEventArgs e)
        {
            ConfigurationLoader.SaveConfiguration<Configuration>(_config);
        }


    }
}
