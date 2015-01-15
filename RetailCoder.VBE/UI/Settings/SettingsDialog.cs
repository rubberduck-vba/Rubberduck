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
        private IConfigurationService _configService;
        private ConfigurationTreeViewControl _treeview;
        private Control _activeControl;

        private TodoSettingController _todoController;
        private TodoListSettingsUserControl _todoView;

        /// <summary>
        ///  Default constructor for GUI Designer. DO NOT USE.
        /// </summary>
        public SettingsDialog()
        {
            InitializeComponent();
        }

        public SettingsDialog(IConfigurationService configService)
            : this()
        {
            _configService = configService;
            _config = _configService.LoadConfiguration();
            _treeview = new ConfigurationTreeViewControl(_config);

            this.splitContainer1.Panel1.Controls.Add(_treeview);
            _treeview.Dock = DockStyle.Fill;
            //todo: switch to controller here 

            var markers = _config.UserSettings.ToDoListSettings.ToDoMarkers.ToList();

            this._todoView = new TodoListSettingsUserControl(markers);

            ActivateControl(this._todoView);

            this._todoController = new TodoSettingController(this._todoView);

            RegisterEvents();
        }

        private void RegisterEvents()
        {
            _treeview.NodeSelected += _treeview_NodeSelected;
        }

        private void _treeview_NodeSelected(object sender, TreeViewEventArgs e)
        {
            Control controlToActivate = null;

            if (e.Node.Text == "Rubberduck")
            {
                return; //do nothing
            }

            if (e.Node.Text == "Todo List")
            {
                //controlToActivate = new TodoListSettingsControl(_config.UserSettings.ToDoListSettings.ToDoMarkers.ToList());
                controlToActivate = this._todoView;
            }

            if (e.Node.Text == "Code Inpsections")
            {

                controlToActivate = new CodeInspectionControl(_config.UserSettings.CodeInspectionSettings.CodeInspections.ToList());
            }

            ActivateControl(controlToActivate);
        }

        private void ActivateControl(Control control)
        {
            control.Dock = DockStyle.Fill;
            this.splitContainer1.Panel2.Controls.Clear();
            this.splitContainer1.Panel2.Controls.Add(control);
            _activeControl = control;
        }

        private void SettingsDialog_FormClosed(object sender, FormClosedEventArgs e)
        {
            SaveConfig();
        }

        private void SaveConfig()
        {
            _config.UserSettings.ToDoListSettings.ToDoMarkers = _todoView.TodoMarkers.ToArray();
            _configService.SaveConfiguration<Configuration>(_config);
        }
    }
}
