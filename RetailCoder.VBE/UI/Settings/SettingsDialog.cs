using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Rubberduck.Inspections;
using Rubberduck.Settings;

namespace Rubberduck.UI.Settings
{
    [ComVisible(true)]
    [Guid(ClassId)]
    [ProgId(ProgId)]
    // ReSharper disable once InconsistentNaming
    public partial class _SettingsDialog : Form
    {
        private const string ClassId = "FB62BEA3-E11A-3C24-9101-AF2E1652AFFC";
        private const string ProgId = "Rubberduck.UI.Settings.SettingsDialog";

        private Configuration _config;
        private IGeneralConfigService _configService;
        private ConfigurationTreeViewControl _treeview;
        private Control _activeControl;

        private TodoSettingPresenter _todoController;
        private TodoListSettingsUserControl _todoView;

        private GeneralSettingsControl _generalSettingsView;

        /// <summary>
        ///  Default constructor for GUI Designer. DO NOT USE.
        /// </summary>
        public _SettingsDialog()
        {
            InitializeComponent();

            OkButton.Click += OkButton_Click;
            CancelButton.Click += CancelButton_Click;

            InitWindow();
        }

        private void InitWindow()
        {
            this.Text = RubberduckUI.Settings_Caption;
            InstructionsLabel.Text = RubberduckUI.SettingsInstructions_GeneralSettings;
            TitleLabel.Text = RubberduckUI.SettingsCaption_GeneralSettings;
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void OkButton_Click(object sender, EventArgs e)
        {
            SaveConfig();
            Close();
        }

        public _SettingsDialog(IGeneralConfigService configService)
            : this()
        {
            _configService = configService;
            _config = _configService.LoadConfiguration();
            _codeInspectionSettings = _config.UserSettings.CodeInspectionSettings.CodeInspections;

            LoadWindow();

            RegisterEvents();
        }

        private void LoadWindow()
        {
            _treeview = new ConfigurationTreeViewControl(_config);

            splitContainer1.Panel1.Controls.Clear();
            splitContainer1.Panel1.Controls.Add(_treeview);
            _treeview.Dock = DockStyle.Fill;

            _generalSettingsView = new GeneralSettingsControl(_config.UserSettings.LanguageSetting, _configService);

            var markers = _config.UserSettings.ToDoListSettings.ToDoMarkers;
            _todoView = new TodoListSettingsUserControl(markers);
            _todoController = new TodoSettingPresenter(_todoView);

            ActivateControl(_generalSettingsView);
        }

        private void RegisterEvents()
        {
            _treeview.NodeSelected += _treeview_NodeSelected;
            _configService.SettingsChanged += _configService_SettingsChanged;
        }

        private void _configService_SettingsChanged(object sender, EventArgs e)
        {
            _config = _configService.LoadConfiguration();

            LoadWindow();
        }

        private readonly IEnumerable<CodeInspectionSetting> _codeInspectionSettings;

        private IEnumerable<CodeInspectionSetting> GetInspectionSettings(CodeInspectionType inspectionType)
        {
            return _codeInspectionSettings.Where(setting => setting.InspectionType == inspectionType);
        }

        private void _treeview_NodeSelected(object sender, TreeViewEventArgs e)
        {
            Control controlToActivate = null;
            if (e.Node == null)
            {
                // a "parent" node is selected. todo: create a page for "parent" nodes.
                return;
            }

            if (e.Node.Text == "Rubberduck")
            {
                TitleLabel.Text = RubberduckUI.SettingsCaption_GeneralSettings;
                InstructionsLabel.Text = RubberduckUI.SettingsInstructions_GeneralSettings;
                ActivateControl(_generalSettingsView);
                return;
            }

            if (e.Node.Text == RubberduckUI.TodoSettings_Caption)
            {
                TitleLabel.Text = RubberduckUI.SettingsCaption_ToDoSettings;
                InstructionsLabel.Text = RubberduckUI.SettingsInstructions_ToDoSettings;
                controlToActivate = _todoView;
            }

            if (e.Node.Parent.Text == RubberduckUI.CodeInspections)
            {
                TitleLabel.Text = RubberduckUI.SettingsCaption_CodeInspections;
                InstructionsLabel.Text = RubberduckUI.SettingsInstructions_CodeInspections;
                var inspectionType = (CodeInspectionType)Enum.Parse(typeof(CodeInspectionType), e.Node.Name);
                var settingGridViewSort = new GridViewSort<CodeInspectionSetting>(RubberduckUI.Name, true);
                controlToActivate = new CodeInspectionSettingsControl(GetInspectionSettings(inspectionType), settingGridViewSort);
            }

            ActivateControl(controlToActivate);
        }

        private void ActivateControl(Control control)
        {
            splitContainer1.Panel2.Controls.Clear();
            splitContainer1.Panel2.Controls.Add(control);
            _activeControl = control;
            try
            {
                _activeControl.Dock = DockStyle.Fill;
            }
            catch { }
        }

        private void SaveConfig()
        {
            var langChanged = !Equals(_config.UserSettings.LanguageSetting, _generalSettingsView.SelectedLanguage);

            _config.UserSettings.LanguageSetting = _generalSettingsView.SelectedLanguage;
            _config.UserSettings.ToDoListSettings.ToDoMarkers = _todoView.TodoMarkers.ToArray();
            // The datagrid view of the CodeInspectionControl seems to keep the config magically in sync, so I don't manually do it here.
            _configService.SaveConfiguration(_config, langChanged);
        }
    }
}
