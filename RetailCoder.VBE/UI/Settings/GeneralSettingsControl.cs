using System;
using System.IO;
using System.Windows.Forms;
using Rubberduck.Settings;

namespace Rubberduck.UI.Settings
{
    public partial class GeneralSettingsControl : UserControl
    {
        private readonly IGeneralConfigService _configService;
        private readonly DisplayLanguageSetting _currentLanguage;

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
            File.Delete(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Rubberduck\\rubberduck.config"));
            var config = _configService.GetDefaultConfiguration();
            _configService.SaveConfiguration(config, !Equals(_currentLanguage, config.UserSettings.LanguageSetting));
        }

        public GeneralSettingsControl(DisplayLanguageSetting displayLanguage, IGeneralConfigService configService)
            : this()
        {
            _configService = configService;
            _currentLanguage = displayLanguage;
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
