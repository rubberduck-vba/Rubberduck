using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Rubberduck.Config;
using Rubberduck.Inspections;

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

            _treeview = new ConfigurationTreeViewControl(_config);

            splitContainer1.Panel1.Controls.Add(_treeview);
            _treeview.Dock = DockStyle.Fill;

            _generalSettingsView = new GeneralSettingsControl(_config.UserSettings.LanguageSetting);

            var markers = _config.UserSettings.ToDoListSettings.ToDoMarkers;
            _todoView = new TodoListSettingsUserControl(markers);
            _todoController = new TodoSettingPresenter(_todoView);

            ActivateControl(_generalSettingsView);
            RegisterEvents();
        }

        private void RegisterEvents()
        {
            _treeview.NodeSelected += _treeview_NodeSelected;
        }

        private readonly IEnumerable<CodeInspectionSetting> _codeInspectionSettings;

        private IEnumerable<CodeInspectionSetting> GetInspectionSettings(CodeInspectionType inspectionType)
        {
            return _codeInspectionSettings.Where(setting => setting.InspectionType == inspectionType);
        }

        private void _treeview_NodeSelected(object sender, TreeViewEventArgs e)
        {
            Control controlToActivate = null;

            if (e.Node.Text == "Rubberduck")
            {
                TitleLabel.Text = RubberduckUI.SettingsCaption_GeneralSettings;
                InstructionsLabel.Text = RubberduckUI.SettingsInstructions_GeneralSettings;
                ActivateControl(_generalSettingsView);
                return;
            }

            if (e.Node.Text == "To-Do Explorer")
            {
                TitleLabel.Text = RubberduckUI.SettingsCaption_ToDoSettings;
                InstructionsLabel.Text = RubberduckUI.SettingsInstructions_ToDoSettings;
                controlToActivate = _todoView;
            }

            if (e.Node.Parent.Text == "Code Inspections")
            {
                TitleLabel.Text = RubberduckUI.SettingsCaption_CodeInspections;
                InstructionsLabel.Text = RubberduckUI.SettingsInstructions_CodeInspections;
                var inspectionType = (CodeInspectionType)Enum.Parse(typeof(CodeInspectionType), e.Node.Text);
                var settingGridViewSort = new GridViewSort<CodeInspectionSetting>(RubberduckUI.Name, true);
                controlToActivate = new CodeInspectionSettingsControl(GetInspectionSettings(inspectionType), settingGridViewSort);
            }

            if (e.Node.Parent.Text == CodeInspectionType.LanguageOpportunities.ToString())
            {
                TitleLabel.Text = RubberduckUI.SettingsCaption_CodeInspections;
                InstructionsLabel.Text = RubberduckUI.SettingsInstructions_CodeInspections;
                var settingGridViewSort = new GridViewSort<CodeInspectionSetting>(RubberduckUI.Name, true);
                controlToActivate = new CodeInspectionSettingsControl(_config.UserSettings.CodeInspectionSettings.CodeInspections.ToList(), settingGridViewSort);
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
            _config.UserSettings.LanguageSetting = _generalSettingsView.SelectedLanguage;
            _config.UserSettings.ToDoListSettings.ToDoMarkers = _todoView.TodoMarkers.ToArray();
            // The datagrid view of the CodeInspectionControl seems to keep the config magically in sync, so I don't manually do it here.
            _configService.SaveConfiguration(_config);
        }
    }
}
