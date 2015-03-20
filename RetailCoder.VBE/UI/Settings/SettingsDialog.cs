using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Rubberduck.Config;

namespace Rubberduck.UI.Settings
{
    [ComVisible(true)]
    public partial class _SettingsDialog : Form
    {
        private Configuration _config;
        private IConfigurationService _configService;
        private ConfigurationTreeViewControl _treeview;
        private Control _activeControl;

        private TodoSettingPresenter _todoController;
        private TodoListSettingsUserControl _todoView;

        /// <summary>
        ///  Default constructor for GUI Designer. DO NOT USE.
        /// </summary>
        public _SettingsDialog()
        {
            InitializeComponent();
        }

        public _SettingsDialog(IConfigurationService configService)
            : this()
        {
            _configService = configService;
            _config = _configService.LoadConfiguration();
            _treeview = new ConfigurationTreeViewControl(_config);

            this.splitContainer1.Panel1.Controls.Add(_treeview);
            _treeview.Dock = DockStyle.Fill;

            var markers = _config.UserSettings.ToDoListSettings.ToDoMarkers.ToList();
            _todoView = new TodoListSettingsUserControl(markers);

            ActivateControl(_todoView);
            _todoController = new TodoSettingPresenter(_todoView);

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
                controlToActivate = _todoView;
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
            MessageBox.Show("Changes to settings will take affect next time the application is started.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void SaveConfig()
        {
            _config.UserSettings.ToDoListSettings.ToDoMarkers = _todoView.TodoMarkers.ToArray();
            // The datagrid view of the CodeInspectionControl seems to keep the config magically in sync, so I don't manually do it here.
            _configService.SaveConfiguration<Configuration>(_config);
        }
    }
}
