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

            OkButton.Click += OkButton_Click;
            CancelButton.Click += CancelButton_Click;
        }

        private void CancelButton_Click(object sender, System.EventArgs e)
        {
            Close();
        }

        private void OkButton_Click(object sender, System.EventArgs e)
        {
            SaveConfig();
            Close();
        }

        public _SettingsDialog(IConfigurationService configService)
            : this()
        {
            _configService = configService;
            _config = _configService.LoadConfiguration();
            _treeview = new ConfigurationTreeViewControl(_config);

            splitContainer1.Panel1.Controls.Add(_treeview);
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
                TitleLabel.Text = RubberduckUI.SettingsCaption_GeneralSettings;
                InstructionsLabel.Text = RubberduckUI.SettingsInstructions_GeneralSettings;
                return; //do nothing
            }

            if (e.Node.Text == "Todo List")
            {
                TitleLabel.Text = RubberduckUI.SettingsCaption_ToDoSettings;
                InstructionsLabel.Text = RubberduckUI.SettingsInstructions_ToDoSettings;
                controlToActivate = _todoView;
            }

            if (e.Node.Text == "Code Inpsections")
            {
                TitleLabel.Text = RubberduckUI.SettingsCaption_CodeInspections;
                InstructionsLabel.Text = RubberduckUI.SettingsInstructions_CodeInspections;
                controlToActivate = new CodeInspectionControl(_config.UserSettings.CodeInspectionSettings.CodeInspections.ToList());
            }

            ActivateControl(controlToActivate);
        }

        private void ActivateControl(Control control)
        {
            control.Dock = DockStyle.Fill;
            splitContainer1.Panel2.Controls.Clear();
            splitContainer1.Panel2.Controls.Add(control);
            _activeControl = control;
        }

        private void SaveConfig()
        {
            _config.UserSettings.ToDoListSettings.ToDoMarkers = _todoView.TodoMarkers.ToArray();
            // The datagrid view of the CodeInspectionControl seems to keep the config magically in sync, so I don't manually do it here.
            _configService.SaveConfiguration(_config);
        }
    }
}
