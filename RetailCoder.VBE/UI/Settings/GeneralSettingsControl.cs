using System;
using System.IO;
using System.Windows.Forms;
using Rubberduck.Settings;

namespace Rubberduck.UI.Settings
{
    public partial class GeneralSettingsControl : UserControl
    {
        private readonly IGeneralConfigService _configService;
        private readonly Configuration _config;

        public GeneralSettingsControl()
        {
            InitializeComponent();
            TitleLabel.Text = RubberduckUI.SettingsCaption_GeneralSettings;
            LanguageLabel.Text = RubberduckUI.Settings_LanguageLabel;
            resetSettings.Text = RubberduckUI.Settings_ResetSettings;

            LoadLanguageList();

            resetSettings.Click += ResetSettingsClick;
        }

        private void ResetSettingsClick(object sender, EventArgs e)
        {
            var confirmReset = MessageBox.Show(RubberduckUI.Settings_ResetSettingsConfirmation, RubberduckUI.Settings_Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
            if (confirmReset == DialogResult.No)
            {
                return;
            }

            ResetSettings();
        }

        private void ResetSettings()
        {
            var oldLang = _config.UserSettings.LanguageSetting;

            File.Delete(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Rubberduck\\rubberduck.config"));
            var config = _configService.GetDefaultConfiguration();
            _configService.SaveConfiguration(config, !oldLang.Equals(config.UserSettings.LanguageSetting));
        }

        public GeneralSettingsControl(DisplayLanguageSetting displayLanguage, Configuration config, IGeneralConfigService configService)
            : this()
        {
            _config = config;
            _configService = configService;
            LanguageList.SelectedItem = displayLanguage;
        }

        public DisplayLanguageSetting SelectedLanguage
        {
            get { return (DisplayLanguageSetting)LanguageList.SelectedItem; }
        }

        private void LoadLanguageList()
        {
            LanguageList.Items.Add(new DisplayLanguageSetting("en-US"));
            LanguageList.Items.Add(new DisplayLanguageSetting("fr-CA"));
            LanguageList.Items.Add(new DisplayLanguageSetting("sv-SV"));

            LanguageList.DisplayMember = "Name";
        }
    }
}
