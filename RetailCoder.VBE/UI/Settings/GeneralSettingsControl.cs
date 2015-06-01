using System.Windows.Forms;
using Rubberduck.Config;

namespace Rubberduck.UI.Settings
{
    public partial class GeneralSettingsControl : UserControl
    {
        public GeneralSettingsControl()
        {
            InitializeComponent();
            TitleLabel.Text = RubberduckUI.SettingsCaption_GeneralSettings;
            InstructionsLabel.Text = RubberduckUI.SettingsInstructions_GeneralSettings;
            LanguageLabel.Text = RubberduckUI.Settings_LanguageLabel;
            resetSettings.Text = RubberduckUI.Settings_ResetSettings;

            LoadLanguageList();

            resetSettings.Click += resetSettings_Click;
        }

        private void resetSettings_Click(object sender, System.EventArgs e)
        {
            var resetSettings = MessageBox.Show(RubberduckUI.Settings_ResetSettingsConfirmation, RubberduckUI.Settings_Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
            if (resetSettings == DialogResult.No)
            {
                return;
            }

            ResetSettings();
        }

        private void ResetSettings()
        {
            throw new System.NotImplementedException();
        }

        public GeneralSettingsControl(DisplayLanguageSetting displayLanguage)
            : this()
        {
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

            LanguageList.DisplayMember = "Name";
        }
    }
}
